#Parameters


$SiteURL = "https://mhud.sharepoint.com/sites/arcce"





  
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



# Get Lock State
$State =@()
$SitesURL | ForEach-Object {
    $State += Get-PnPTenantSite -Identity $_ | Select-Object Description, URL, LockState   
    }
$State | Format-Table -AutoSize -Wrap

