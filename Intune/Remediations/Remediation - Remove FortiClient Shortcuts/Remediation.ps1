try {
    $shortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\FortiClient"
    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $publicDesktopPath = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
    
    # Remove FortiClient Start Menu folder
    if (Test-Path -Path $shortcutPath) {
        Remove-Item -Path $shortcutPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Removed FortiClient Start Menu shortcuts folder."
    }
    
    # Remove FortiClient shortcuts from both user and public desktop
    $fortiDesktopShortcuts = Get-ChildItem -Path $desktopPath, $publicDesktopPath -Filter "FortiClient*.lnk" -ErrorAction SilentlyContinue
    if ($fortiDesktopShortcuts) {
        $fortiDesktopShortcuts | ForEach-Object {
            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Output "Removed desktop shortcut: $($_.Name)"
        }
    }
    
    Write-Output "Cleanup completed successfully."
    Exit 0
} catch {
    $errMsg = $_.Exception.Message
    Write-Output $errMsg
    Exit 1
}