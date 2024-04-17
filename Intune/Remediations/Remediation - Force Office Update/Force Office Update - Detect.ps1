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
        Exit 1
    } else {
        Write-Host "Current version is up to date: $currentVersion"
        Exit 0
    }
}
catch {
    $errMsg = $_.exeption.message
    Write-Output $errMsg
    Exit 0
}
