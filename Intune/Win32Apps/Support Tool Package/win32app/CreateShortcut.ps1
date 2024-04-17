$scriptPath = "C:\Code\Github\HUD-Microsoft-365\HUD Support Scripts\InvokeTool.ps1"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "Digital Support Admin Shell.lnk"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = "powershell.exe"
$Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
$Shortcut.Save()