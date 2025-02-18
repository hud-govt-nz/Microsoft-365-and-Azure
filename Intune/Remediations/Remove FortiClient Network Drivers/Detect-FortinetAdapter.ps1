#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

try {
    # Check for Fortinet adapters
    $adapters = Get-NetAdapter | Where-Object { 
        $_.DriverDescription -like "*Fortinet SSL VPN*" -or 
        $_.DriverDescription -like "*Fortinet Virtual Ethernet Adapter*" 
    }

    # Check for Fortinet PnP devices
    $devices = Get-PnpDevice | Where-Object { 
        ($_.FriendlyName -like "*Fortinet SSL VPN*" -or 
         $_.FriendlyName -like "*Fortinet Virtual Ethernet Adapter*")
    }

    # Check for Fortinet drivers
    $drivers = Get-WindowsDriver -Online | Where-Object {
        $_.OriginalFileName -like "*Fortinet*" -or
        $_.ProviderName -like "*Fortinet*"
    }

    if ($null -eq $adapters -and $null -eq $devices -and $null -eq $drivers) {
        Write-Verbose "No Fortinet Virtual Ethernet components found - compliant"
        exit 0
    } else {
        $components = @()
        if ($adapters) { $components += "Adapters: $($adapters.Count)" }
        if ($devices) { $components += "Devices: $($devices.Count)" }
        if ($drivers) { $components += "Drivers: $($drivers.Count)" }
        
        Write-Verbose "Found Fortinet components: $($components -join ', ') - non-compliant"
        exit 1
    }
}
catch {
    Write-Error "Error in detection script: $_"
    exit 1
}