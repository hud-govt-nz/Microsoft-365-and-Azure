function Test-OneDriveProcess {
    return [bool](Get-Process "OneDrive" -ErrorAction SilentlyContinue)
}

# Ensure running in user context
if ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM") {
    Write-Error "Running as SYSTEM, this script should run in user context!"
    exit 1
}

# OneDrive detection
if (Test-OneDriveProcess) {
    Write-Output "OneDrive is running, no action required."
    exit 0
} else {
    Write-Output "OneDrive is not running, remediation is needed."
    exit 1
}
