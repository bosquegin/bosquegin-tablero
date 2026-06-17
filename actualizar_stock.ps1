# actualizar_stock.ps1
# Lee Stock_consolidado_por_deposito_y_dia.xlsx y actualiza STOCK_CIERRE_MES
# en bosquegin_dashboard.html con el stock real KLOZER+OFI por mes.
# Uso: .\actualizar_stock.ps1

$consolidadoPath = "$PSScriptRoot\Data\Inventario\Stock_consolidado_por_deposito_y_dia.xlsx"
$dashboardPath   = "$PSScriptRoot\bosquegin_dashboard.html"

Write-Host "=== Actualizando stock KLOZER+OFI ===" -ForegroundColor Cyan
Write-Host "Fuente: $consolidadoPath"

if (-not (Test-Path $consolidadoPath)) {
  Write-Error "No se encontro el archivo consolidado: $consolidadoPath"; exit 1
}

# --- Leer Excel via COM ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($consolidadoPath)
$sh = $wb.Sheets.Item(1)
$lastRow = $sh.UsedRange.Rows.Count
Write-Host "Filas en consolidado: $lastRow"

# Leer filas KLOZER+OFI con codigo numerico y stock > 0
$rows = @()
for ($r = 1; $r -le $lastRow; $r++) {
  $dep = $sh.Cells.Item($r, 2).Text
  if ($dep -ne "KLOZER" -and $dep -ne "OFI") { continue }
  $dt  = $sh.Cells.Item($r, 1).Text   # "YYYY-MM-DD"
  $cod = $sh.Cells.Item($r, 4).Text
  $qty = $sh.Cells.Item($r, 6).Value2
  if ($cod -match "^\d{5,6}$" -and $qty -gt 0) {
    $rows += [PSCustomObject]@{ Fecha = $dt; Cod = $cod; Qty = [int]$qty }
  }
}
$wb.Close($false); $xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
Write-Host "Filas KLOZER+OFI validas: $($rows.Count)"

# --- Encontrar ultimo dia disponible por mes ---
$lastDayByMonth = @{}
foreach ($row in $rows) {
  $parts = $row.Fecha -split "-"
  $key = "$([int]$parts[0])_$([int]$parts[1])"
  if (-not $lastDayByMonth[$key] -or $row.Fecha -gt $lastDayByMonth[$key]) {
    $lastDayByMonth[$key] = $row.Fecha
  }
}

Write-Host "`nMeses encontrados:"
$lastDayByMonth.Keys | Sort-Object | ForEach-Object {
  Write-Host "  $_ -> $($lastDayByMonth[$_])"
}

# --- Sumar stock por codigo en el ultimo dia de cada mes ---
$stockByMonth = @{}
foreach ($row in $rows) {
  $parts = $row.Fecha -split "-"
  $key = "$([int]$parts[0])_$([int]$parts[1])"
  if ($row.Fecha -ne $lastDayByMonth[$key]) { continue }
  if (-not $stockByMonth[$key]) { $stockByMonth[$key] = @{} }
  if (-not $stockByMonth[$key][$row.Cod]) { $stockByMonth[$key][$row.Cod] = 0 }
  $stockByMonth[$key][$row.Cod] += $row.Qty
}

# --- Generar bloque JS para cada mes ---
function Build-MonthBlock($key, $date, $data) {
  $lines = @()
  $lines += "  // $key`: Stock_consolidado_por_deposito_y_dia.xlsx — KLOZER+OFI — $date"
  $lines += "  `"$key`": {"

  $sorted = $data.Keys | Sort-Object
  $entries = $sorted | ForEach-Object { "`"$_`": $($data[$_])" }

  # Agrupar en lineas de hasta 4 entradas para legibilidad
  $chunk = 4
  for ($i = 0; $i -lt $entries.Count; $i += $chunk) {
    $slice = $entries[$i..([Math]::Min($i + $chunk - 1, $entries.Count - 1))]
    $comma = if ($i + $chunk -lt $entries.Count) { "," } else { "" }
    $lines += "    $($slice -join ', ')$comma"
  }
  $lines += "  }"
  return $lines -join "`n"
}

# --- Construir el bloque completo de meses 2026 ---
$monthKeys = $stockByMonth.Keys | Sort-Object
$blocks = @()
foreach ($key in $monthKeys) {
  $block = Build-MonthBlock $key $lastDayByMonth[$key] $stockByMonth[$key]
  $blocks += $block
}
$newBlockContent = $blocks -join ",`n"

# --- Leer dashboard y reemplazar seccion 2026 ---
$html = Get-Content $dashboardPath -Raw -Encoding UTF8

# Patron: desde el primer "2026_" hasta el cierre de la ultima entrada 2026 antes de };
# Buscamos el inicio del primer bloque 2026 y el final del ultimo
$firstYearKey = ($monthKeys | Where-Object { $_ -like "2026_*" } | Sort-Object | Select-Object -First 1)
$lastYearKey  = ($monthKeys | Where-Object { $_ -like "2026_*" } | Sort-Object | Select-Object -Last 1)

if (-not $firstYearKey) {
  Write-Error "No se encontraron meses 2026 en el consolidado."; exit 1
}

# Encontrar el bloque en el HTML usando regex
# El patron captura desde el comentario/clave del primer mes 2026 hasta el cierre del ultimo
$pattern = '(?s)(  // 2026_1[^\n]*\n  "2026_1": \{.*?"' + $lastYearKey + '": \{.*?\})'
if ($html -notmatch $pattern) {
  # Intento alternativo sin comentario previo
  $pattern = '(?s)("2026_1": \{.*?"' + $lastYearKey + '": \{.*?\})'
}

if ($html -match $pattern) {
  $oldBlock = $Matches[1]
  $html = $html.Replace($oldBlock, $newBlockContent)
  $html | Set-Content $dashboardPath -Encoding UTF8 -NoNewline
  Write-Host "`nDashboard actualizado correctamente." -ForegroundColor Green
  Write-Host "Meses actualizados: $($monthKeys -join ', ')"
} else {
  Write-Error "No se pudo encontrar el bloque 2026 en el dashboard. Actualizacion manual requerida."
  Write-Host "`nBloque generado (copiar manualmente si es necesario):"
  Write-Host $newBlockContent
}
