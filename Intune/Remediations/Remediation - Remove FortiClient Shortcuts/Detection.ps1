try {
    $shortcutPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\FortiClient"
    $desktopPath = [System.Environment]::GetFolderPath('Desktop')
    $publicDesktopPath = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
    
    $needsRemediation = $false
    
    # Check if FortiClient Start Menu folder exists
    if (Test-Path -Path $shortcutPath) {
        Write-Output "FortiClient Start Menu shortcuts folder found."
        $needsRemediation = $true
    }
    
    # Check for FortiClient shortcuts on both user and public desktop
    $fortiDesktopShortcuts = Get-ChildItem -Path $desktopPath, $publicDesktopPath -Filter "FortiClient*.lnk" -ErrorAction SilentlyContinue
    if ($fortiDesktopShortcuts) {
        Write-Output "FortiClient desktop shortcuts found."
        $needsRemediation = $true
    }
    
    if ($needsRemediation) {
        Exit 1
    } else {
        Write-Output "No FortiClient shortcuts found. No remediation needed."
        Exit 0
    }
} catch {
    $errMsg = $_.Exception.Message
    Write-Output $errMsg
    Exit 1
}