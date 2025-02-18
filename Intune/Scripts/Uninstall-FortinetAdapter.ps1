#Requires -RunAsAdministrator

# Script to uninstall Fortinet SSL VPN Virtual Ethernet Adapter
$ErrorActionPreference = 'Stop'

function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host $logMessage
}

try {
    Write-Log "Starting Fortinet Virtual Ethernet Adapter uninstallation..."
    
    # Get all matching network adapters
    $adapters = Get-NetAdapter | Where-Object { 
        $_.DriverDescription -like "*Fortinet SSL VPN*" -or 
        $_.DriverDescription -like "*Fortinet Virtual Ethernet Adapter*" 
    }
    
    if ($null -eq $adapters) {
        Write-Log "Fortinet Virtual Ethernet Adapter not found."
        exit 0
    }

    foreach ($adapter in $adapters) {
        Write-Log "Processing adapter: $($adapter.Name)"

        # Disable the adapter first
        Write-Log "Disabling adapter..."
        Disable-NetAdapter -Name $adapter.Name -Confirm:$false

        # Get the corresponding PnP device
        $device = Get-PnpDevice | Where-Object { 
            ($_.FriendlyName -like "*Fortinet SSL VPN*" -or 
             $_.FriendlyName -like "*Fortinet Virtual Ethernet Adapter*") 
        }

        if ($device) {
            Write-Log "Uninstalling device with ID: $($device.InstanceId)"
            
            # Remove the device
            & pnputil /remove-device "$($device.InstanceId)" /force
            if ($LASTEXITCODE -eq 0) {
                Write-Log "Successfully removed device $($adapter.Name)"
            } else {
                Write-Log "Warning: Device removal returned exit code $LASTEXITCODE for $($adapter.Name)"
            }

            # Get and remove the actual driver
            $driverInfo = Get-WindowsDriver -Online | Where-Object {
                $_.OriginalFileName -like "*Fortinet*" -or
                $_.ProviderName -like "*Fortinet*"
            }

            if ($driverInfo) {
                foreach ($driver in $driverInfo) {
                    Write-Log "Removing driver: $($driver.Driver)"
                    & pnputil /delete-driver $driver.Driver /force /uninstall
                    if ($LASTEXITCODE -eq 0) {
                        Write-Log "Successfully uninstalled driver $($driver.Driver)"
                    } else {
                        Write-Log "Warning: Driver removal returned exit code $LASTEXITCODE for $($driver.Driver)"
                    }
                }
            } else {
                Write-Log "No Fortinet drivers found to remove"
            }
        }
        else {
            Write-Log "Could not find matching PnP device for adapter $($adapter.Name)"
        }
    }

    Write-Log "Uninstallation completed successfully."
    Write-Log "Note: A system restart may be required to complete the uninstallation."
}
catch {
    Write-Log "Error occurred during uninstallation: $_"
    exit 1
}