param (
    [switch]$IncludeOneDriveSites
)

Clear-Host
Write-host "## SharePoint Online: Bulk Remove Files or Folders ##" -ForegroundColor Yellow

$AdminSiteURL = "https://mhud-admin.sharepoint.com"

#Requires -Modules PNP.Powershell
# Connect to PnP PowerShell
try {
    $env:PNPPOWERSHELL_UPDATECHECK = "Off"
    Connect-PnPOnline -Url $AdminSiteURL `
        -ClientId $env:DigitalSupportAppID `
        -Tenant 'mhud.onmicrosoft.com' `
        -Thumbprint $env:DigitalSupportCertificateThumbprint
    Write-Host "Connected to SharePoint Administration Portal" -ForegroundColor Green
} catch {
    Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
    exit 1
}

# Define the mapping from the key to the descriptive text
$TemplateMappings = @{
    'APPCATALOG#0'               = 'App Catalog Site'
    'BDR#0'                      = 'Document Center'
    'BICenterSite#0'             = 'Business Intelligence Center'
    'BLANKINTERNET#0'            = 'Publishing Site'
    'BLANKINTERNETCONTAINER#0'   = 'Publishing Portal'
    'COMMUNITY#0'                = 'Community Site'
    'COMMUNITYPORTAL#0'          = 'Community Portal'
    'DEV#0'                      = 'Developer Site'
    'EHS#1'                      = 'Team Site - SharePoint Online configuration'
    'ENTERWIKI#0'                = 'Enterprise Wiki'
    'GROUP#0'                    = 'Team site'
    'OFFILE#1'                   = 'Records Center'
    'POINTPUBLISHINGHUB#0'       = 'PointPublishing Hub'
    'POINTPUBLISHINGPERSONAL#0'  = 'Personal blog'
    'POINTPUBLISHINGTOPIC#0'     = 'PointPublishing Topic'
    'PRODUCTCATALOG#0'           = 'Product Catalog'
    'PROJECTSITE#0'              = 'Project Site'
    'PWA#0'                      = 'Project Web App Site'
    'RedirectSite#0'             = 'Redirect Site'
    'REVIEWCTR#0'                = 'Review Center'
    'SITEPAGEPUBLISHING#0'       = 'Communication site'
    'SPSMSITEHOST#0'             = 'My Site Host'
    'SRCHCEN#0'                  = 'Enterprise Search Center'
    'SRCHCENTERLITE#0'           = 'Basic Search Center'
    'STS#0'                      = 'Team site (classic experience)'
    'STS#3'                      = 'Team site (no Microsoft 365 group)'
    'TEAMCHANNEL#0'              = 'Team channel'
    'TEAMCHANNEL#1'              = 'Team channel'
    'visprus#0'                  = 'Visio Process Repository'
}

# Obtain Site Data
$Sites = @()

if ($IncludeOneDriveSites) {
    Get-PnPTenantSite -Detailed -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"
    | ForEach-Object {
        $item = $_ | Select-Object Title, Url, Owner, @{Name='Template'; Expression={$TemplateMappings[$_.Template]} }
        $Sites += $item
    } | Out-Null
} else {
    Get-PnPTenantSite -Detailed | ForEach-Object {
        $item = $_ | Select-Object Title, Url, Owner, @{Name='Template'; Expression={$TemplateMappings[$_.Template]} }
        $Sites += $item
    } | Out-Null
}

$SelectedSite = $Sites | Out-GridView -Title "Sites" -PassThru
$SiteURL = $SelectedSite.Url

# Connect to the selected site
Write-Host "Connecting to the selected site: $($SelectedSite.Title)" -ForegroundColor Yellow
Connect-PnPOnline `
    -Url $SiteURL `
    -ClientId $env:DigitalSupportAppID `
    -Tenant 'mhud.onmicrosoft.com' `
    -Thumbprint $env:DigitalSupportCertificateThumbprint

# Retrieve and select the list
if ($SelectedSite.Url -like "*-my.sharepoint.com*") {
    $rootUrl = "https://mhud-my.sharepoint.com"
} else {
    $rootUrl = "https://mhud.sharepoint.com"
}
$Lists = Get-PnPList
$SelectedList = $Lists | Out-GridView -Title "Select a List" -PassThru
$ListName = $SelectedList.Title
$ListURL = $SelectedList.DefaultViewUrl

# Open the site URL in the default web browser
Start-Process ("$rootUrl" + "$ListURL").ToString()

Write-Host "Opening the site library in the default web browser..." -ForegroundColor Yellow

#Config Variables

# Prompt the user to enter the folder's server relative URL
Write-Host "Example: /sites/YourSiteName/Shared Documents/YourFolderName"

$FolderServerRelativeURL = Read-Host "Enter the Folder Server Relative URL"

# Confirm deletion
$confirmation = Read-Host "Are you sure you want to delete all files in the folder '$FolderServerRelativeURL'? Type 'Yes' to confirm"

if ($confirmation -eq 'Yes') {
    Try {
        #Get All Items from Folder in Batch
        $ListItems = Get-PnPListItem -List $ListName -FolderServerRelativeUrl $FolderServerRelativeURL -PageSize 2000 | Sort-Object ID -Descending
       
        #Powershell to delete all files from a folder
        ForEach ($Item in $ListItems) {
            Remove-PnPListItem -List $ListName -Identity $Item.Id -Recycle -Force | Out-Null
            Write-host "Removed File:" $Item.FieldValues.FileRef -ForegroundColor DarkBlue   
        }
    }
    Catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
} else {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
}