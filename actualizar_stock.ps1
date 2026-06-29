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

# --- Leer meses existentes de data_stock_cierre.js (preservar los que no están en el Excel) ---
$scmPath = Join-Path $PSScriptRoot "data_stock_cierre.js"
$existingMonths = @{}  # key → raw JSON block string

if (Test-Path $scmPath) {
  $existingJs = [System.IO.File]::ReadAllText($scmPath, [System.Text.Encoding]::UTF8)
  # Extract each "YYYY_M": { ... } block
  $rx = [regex]'"(\d{4}_\d{1,2})"\s*:\s*(\{[^}]*\})'
  foreach ($m in $rx.Matches($existingJs)) {
    $existingMonths[$m.Groups[1].Value] = $m.Groups[2].Value
  }
  Write-Host "Meses existentes en data_stock_cierre.js: $($existingMonths.Keys | Sort-Object)"
}

# --- Generar bloque JS para cada mes ---
function Build-MonthBlock($key, $date, $data) {
  $lines = @()
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

# --- Fusionar: meses existentes + meses nuevos del Excel ---
# Los meses del Excel reemplazan los existentes; el resto se conserva
foreach ($key in $stockByMonth.Keys) {
  $existingMonths[$key] = $null  # marcar para reemplazar con datos frescos del Excel
}

# Construir bloques ordenados
$allKeys = ($existingMonths.Keys + $stockByMonth.Keys) | Select-Object -Unique | Sort-Object {
  $parts = $_ -split "_"; [int]$parts[0] * 100 + [int]$parts[1]
}
$blocks = @()
foreach ($key in $allKeys) {
  if ($stockByMonth.ContainsKey($key)) {
    # Datos frescos del Excel
    $blocks += Build-MonthBlock $key $lastDayByMonth[$key] $stockByMonth[$key]
  } elseif ($existingMonths[$key]) {
    # Mes existente (ej: 2025) que no está en el Excel — preservar
    $blocks += "  `"$key`": $($existingMonths[$key])"
  }
}
$newBlockContent = $blocks -join ",`n"

# Escribir data_stock_cierre.js
$scmJs = "window.STOCK_CIERRE_MES={`n$newBlockContent`n};"
[System.IO.File]::WriteAllText($scmPath, $scmJs, [System.Text.Encoding]::UTF8)
Write-Host "`ndata_stock_cierre.js escrito correctamente." -ForegroundColor Green
Write-Host "Meses en archivo: $($allKeys -join ', ')"
