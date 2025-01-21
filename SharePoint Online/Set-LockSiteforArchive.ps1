#Parameters
$TenantAdminURL = "https://mhud-admin.sharepoint.com"

$SiteURL = "https://mhud.sharepoint.com/sites/arcce"


$SitesURL =@(
            "https://mhud.sharepoint.com/sites/hud201991714557", 
            "https://mhud.sharepoint.com/sites/arcce"
            "https://mhud.sharepoint.com/sites/arcem",
            "https://mhud.sharepoint.com/sites/arcgs",
            "https://mhud.sharepoint.com/sites/arclp",
            "https://mhud.sharepoint.com/sites/arcpa",
            "https://mhud.sharepoint.com/sites/arcxb", 
            "https://mhud.sharepoint.com/sites/hud20201815248",
            "https://mhud.sharepoint.com/sites/arcmf",
            "https://mhud.sharepoint.com/sites/arcmh",
            "https://mhud.sharepoint.com/sites/arcmr",
            "https://mhud.sharepoint.com/sites/arcms",
            "https://mhud.sharepoint.com/sites/arcmu",
            "https://mhud.sharepoint.com/sites/arctr"
            )


  
Try {
    #Connect to Admin Center
    Connect-PnPOnline -Url $TenantAdminURL -Interactive
     
    #Get Lock Status of the site
    Get-PnPTenantSite -Identity $SiteURL | Select-Object URL, LockState
}
Catch {
    Write-Host -f Red "Error:"$_.Exception.Message
}


# https://www.sharepointdiary.com/2017/08/how-to-lock-site-collection-in-sharepoint-online.html#ixzz8TXLj6O3X
# https://www.sharepointdiary.com/2019/02/set-sharepoint-online-site-to-read-only-using-powershell.html

# Set lock state to ReadOnly
$SitesURL | ForEach-Object {
    #Set-PnPSite -Identity $_ -LockState ReadOnly
}

# Get Lock State
$State =@()
$SitesURL | ForEach-Object {
    $State += Get-PnPTenantSite -Identity $_ | Select-Object Description, URL, LockState   
    }
$State | Format-Table -AutoSize -Wrap

# Unlock site
$SitesURL | ForEach-Object {
    Set-PnPSite -Identity $_ -LockState Unlock
}