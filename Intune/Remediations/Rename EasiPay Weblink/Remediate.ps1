$ShortcutPath = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\EasiPay.url'
if (Test-Path $ShortcutPath) {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.Save()
    Rename-Item $ShortcutPath -NewName 'HUD Pay.Url'
}
