#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

function Write-Log {
    param($Message)
    Write-Verbose $Message
}

try {
    Write-Log "Starting Fortinet Virtual Ethernet Adapter remediation..."
    $remediationNeeded = $false
    
    # Handle network adapters
    $adapters = Get-NetAdapter | Where-Object { 
        $_.DriverDescription -like "*Fortinet SSL VPN*" -or 
        $_.DriverDescription -like "*Fortinet Virtual Ethernet Adapter*" 
    }
    
    if ($adapters) {
        $remediationNeeded = $true
        foreach ($adapter in $adapters) {
            Write-Log "Disabling adapter: $($adapter.Name)"
            Disable-NetAdapter -Name $adapter.Name -Confirm:$false
        }
    }

    # Handle PnP devices
    $devices = Get-PnpDevice | Where-Object { 
        ($_.FriendlyName -like "*Fortinet SSL VPN*" -or 
         $_.FriendlyName -like "*Fortinet Virtual Ethernet Adapter*")
    }

    if ($devices) {
        $remediationNeeded = $true
        foreach ($device in $devices) {
            Write-Log "Removing device: $($device.FriendlyName)"
            & pnputil /remove-device "$($device.InstanceId)" /force
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Successfully removed device $($device.FriendlyName)"
            } else {
                Write-Log "Warning: Device removal returned exit code $LASTEXITCODE for $($device.FriendlyName)"
            }
        }
    }

    # Handle drivers
    $drivers = Get-WindowsDriver -Online | Where-Object {
        $_.OriginalFileName -like "*Fortinet*" -or
        $_.ProviderName -like "*Fortinet*"
    }

    if ($drivers) {
        $remediationNeeded = $true
        foreach ($driver in $drivers) {
            Write-Log "Removing driver: $($driver.Driver)"
            & pnputil /delete-driver $driver.Driver /force /uninstall
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Successfully removed driver $($driver.Driver)"
            } else {
                Write-Log "Warning: Driver removal returned exit code $LASTEXITCODE for $($driver.Driver)"
            }
        }
    }

    if ($remediationNeeded) {
        Write-Log "Remediation actions completed. A system restart may be required."
        exit 0  # Success
    } else {
        Write-Log "No Fortinet components found requiring remediation."
        exit 0  # Nothing to do
    }
}
catch {
    Write-Error "Error in remediation script: $_"
    exit 1
}