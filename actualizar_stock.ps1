# actualizar_stock.ps1
# Lee Stock_consolidado_por_deposito_y_dia.xlsx y escribe data_stock_cierre.js
# Uso: .\actualizar_stock.ps1

$consolidadoPath = "$PSScriptRoot\Data\Inventario\Stock_consolidado_por_deposito_y_dia.xlsx"

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


# Escribir data_stock_cierre.js (ya no se embebe en el HTML)
$scmJs = "window.STOCK_CIERRE_MES={`n$newBlockContent`n};"
$scmPath = Join-Path $PSScriptRoot "data_stock_cierre.js"
$scmJs | Set-Content $scmPath -Encoding UTF8 -NoNewline
Write-Host "`ndata_stock_cierre.js escrito correctamente." -ForegroundColor Green
Write-Host "Meses actualizados: $($monthKeys -join ', ')"
