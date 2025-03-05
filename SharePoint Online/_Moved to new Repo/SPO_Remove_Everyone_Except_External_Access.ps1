### ALL SITES

$AdminURL = "https://mhud-admin.sharepoint.com"

$env:PNPPOWERSHELL_UPDATECHECK = "Off"
Connect-PnPOnline -Url $AdminURL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
 
# Get all SharePoint Online sites
$AllSites =  Get-PnPTenantSite | Where-Object -Property Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")
 
#Loop through each site collection
ForEach($Site in $AllSites)
{
    Write-host -f Magenta "Processing site:" $Site.URL        
 
    #Connect to the Site
    Connect-PnPOnline -URL $Site.URL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    
    #Check if the site contains any permissions (Direct/Group Membershipo) to "Everyone except external users"
    $EEEUsers = Get-PnPUser  | Where-Object {$_.Title -eq "Everyone except external users"}
 
    If($EEEUsers)
    {
        Write-host -f Yellow -NoNewline "`tFound the 'Everyone except external users' group on the site! "
     
        #Remove user from the site    
        Remove-PnPUser -Identity "Everyone except external users" -Force
        Write-host -f Green "Removed!"
    }
}

### SPECIFIC SITE

$SiteURL = "https://mhud.sharepoint.com/sites/dms-ministerial"

$env:PNPPOWERSHELL_UPDATECHECK = "Off"
Connect-PnPOnline -Url $SiteURL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint


#Get the Tenant ID
$TenantID = Get-PnPTenantId
$SearchGroupID = "spo-grid-all-users/$TenantID" #Everyone except external users
$EEEUsersID = "c:0-.f|rolemanager|$SearchGroupID"
 
#Get all site groups
$allGroups = Get-PnPSiteGroup -Site $SiteURL
$totalGroups = $allGroups.Count
$Groups = @()
$i = 0

foreach ($group in $allGroups) {
    $i++
    $percent = ($i / $totalGroups) * 100
    Write-Progress -Activity "Scanning groups" -Status "Processing group: $($group.Title)" -PercentComplete $percent

    if ($group.Users -contains $SearchGroupID) {
        $Groups += $group
    }
}

If($Groups) {
    Write-host -f Yellow -NoNewline "Found the Group under: " ($Groups.Title -join "; ")
    #Remove from the Group(s)
    $Groups | ForEach-Object { Remove-PnPGroupMember -LoginName $EEEUsersID -Identity $_.Title }
    Write-host -f Green "`tRemoved from the Group(s)!"
}