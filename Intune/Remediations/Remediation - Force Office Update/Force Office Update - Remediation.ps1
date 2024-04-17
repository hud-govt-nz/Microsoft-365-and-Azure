# Define variables
$oldVersion = "15.0.1234.5678"
$appName = "Microsoft 365 Apps for enterprise - en-us"

# Check current version
$Installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object DisplayName, DisplayVersion
$Installed += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion
$Result = $Installed | Where-Object { $_.DisplayName -ne $null } | Where-Object {$_.DisplayName -match $appName }

$currentVersion = $result.DisplayVersion

try {
    if ($currentVersion -lt $oldVersion) {
        Write-Host "Old version detected: $currentVersion"
        
        # Force update
        Write-Host "Forcing update to version $newVersion..."
        Start-Process -WindowStyle hidden -FilePath "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" -ArgumentList "/update user updatepromptuser=false forceappshutdown=false displaylevel=false" -Wait
    
        # Check updated version
        $UpdatedInstall = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |  Select-Object DisplayName, DisplayVersion
        $UpdatedInstall += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion
        $Result = $UpdatedInstall | Where-Object { $_.DisplayName -ne $null } | Where-Object {$_.DisplayName -match $appName }

        $updatedVersion = $result.DisplayVersion
        
        if ($updatedVersion -eq $newVersion) {
            Write-Host "Update successful!"
        } else {
            Write-Host "Update failed. Current version: $updatedVersion"
        }
    } else {
        Write-Host "Current version is up to date: $currentVersion"
    }
}
catch {
    $errMsg = $_.exeption.message
    Write-Output $errMsg
    Exit 0
}
