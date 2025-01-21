# Set error action preference to stop  
$ErrorActionPreference = 'Stop'  
  
try {  
    # Obtain Wi-Fi Network Adapter  
    $WIFIProfile = Get-NetAdapter | Where-Object { $_.Name -like "*wi*" }  
    if (-not $WIFIProfile) {  
        throw "No Wi-Fi network adapter found."  
    }  
    $Interface = $WIFIProfile.Name  
  
    # Set Network Interface Profile Order  
    netsh wlan set profileorder name="HUD-Corporate" interface="$Interface" priority=1  
    netsh wlan set profileorder name="SA_CORP" interface="$Interface" priority=2  
    netsh wlan set profileorder name="HUD-CORP" interface="$Interface" priority=3  
  
    # Set network profiles to their desired connection modes  
    netsh wlan set profileparameter name="HUD-Corporate" connectionmode=auto interface="$Interface"  
    netsh wlan set profileparameter name="SA_CORP" connectionmode=auto interface="$Interface"  
    netsh wlan set profileparameter name="HUD-CORP" connectionmode=auto interface="$Interface"  
}  
catch {  
    # Handle errors  
    Write-Error "An error occurred: $_"  
    exit 1  
}  