# setup_brave_cdp.ps1
# Agrega --remote-debugging-port=9222 a los accesos directos de Brave.
# Se ejecuta automaticamente al actualizar el tablero.

$flag = "--remote-debugging-port=9222"
$ws   = New-Object -ComObject WScript.Shell

$candidates = @(
    "$env:USERPROFILE\Desktop\Brave Browser.lnk",
    "$env:USERPROFILE\Desktop\Brave.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Brave Browser.lnk",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Brave.lnk",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Brave Browser.lnk",
    "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Brave.lnk",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Brave Browser.lnk",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Brave.lnk",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Brave Browser.lnk",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Brave.lnk"
)

# Buscar tambien en el escritorio por si tiene otro nombre
$desktop = "$env:USERPROFILE\Desktop"
Get-ChildItem $desktop -Filter "*.lnk" | Where-Object { $_.Name -like "*Brave*" } | ForEach-Object {
    if ($candidates -notcontains $_.FullName) { $candidates += $_.FullName }
}

$modified = 0
$already  = 0

foreach ($path in $candidates) {
    if (-not (Test-Path $path)) { continue }
    try {
        $sc = $ws.CreateShortcut($path)
        if ($sc.Arguments -like "*remote-debugging-port*") {
            $already++
        } else {
            $sc.Arguments = ($sc.Arguments + " " + $flag).Trim()
            $sc.Save()
            $modified++
        }
    } catch { }
}
