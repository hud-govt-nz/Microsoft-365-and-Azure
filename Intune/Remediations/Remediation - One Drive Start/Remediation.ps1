param(
    [int]$OneDriveConfigWaitTime = 600,
    [int]$OneDriveRestartWaitTime = 900
)

function Stop-OneDriveProcess {
    Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue
}

function Start-OneDriveProcess {
    Start-Process "$env:SYSTEMROOT\System32\OneDriveSetup.exe" -ArgumentList "/thfirstsetup"
}

# Ensure running in user context
if ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -eq "NT AUTHORITY\SYSTEM") {
    Write-Error "Running as SYSTEM, this script should run in user context!"
    exit 1
}

# Attempt to start OneDrive
try {
    Start-OneDriveProcess
    Start-Sleep -Seconds $OneDriveConfigWaitTime
} catch {
    Write-Error "Failed to start OneDrive. Exiting script."
    exit 1
}

# Validate remediation
if (Test-OneDriveProcess) {
    Write-Output "Successfully started OneDrive."
    exit 0
} else {
    Write-Output "Failed to start OneDrive after first attempt. Trying once more."

    # Try again to remediate
    try {
        Stop-OneDriveProcess
        Start-Sleep -Seconds $OneDriveRestartWaitTime
        Start-OneDriveProcess
    } catch {
        Write-Error "Failed to restart OneDrive. Exiting script."
        exit 1
    }

    if (Test-OneDriveProcess) {
        Write-Output "Successfully started OneDrive after second attempt."
        exit 0
    } else {
        Write-Error "Failed to start OneDrive after second attempt. Exiting script."
        exit 1
    }
}
