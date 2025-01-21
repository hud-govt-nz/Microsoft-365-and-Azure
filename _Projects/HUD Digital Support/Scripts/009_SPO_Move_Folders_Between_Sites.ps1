Clear-Host
Write-host "## SharePoint Online: Move Folders Between Sites/Libraries ##" -ForegroundColor Yellow

$SiteURL = "https://mhud-admin.sharepoint.com"

#Requires -Modules PNP.Powershell
# Connect to PnP PowerShell
try {
    Write-Host "Connecting to PnP PowerShell..."
    Connect-PnPOnline -Url $SiteURL `
                      -ClientId $env:DigitalSupportAppID `
                      -Tenant 'mhud.onmicrosoft.com' `
                      -Thumbprint $env:DigitalSupportCertificateThumbprint
    Write-Host "Connected" -ForegroundColor Green
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
$Sites =@()

Get-PnPTenantSite -Detailed | ForEach-Object {
    $item = $_ | Select-Object Title, Url, Owner,
    @{Name='Template'; Expression={$TemplateMappings[$_.Template]}}
    
    $Sites += $item
    } | Out-Null

$SelectedSite = $Sites | Out-GridView -Title "SharePoint Sites" -PassThru

$SiteURL = $SelectedSite

#Connect to PnP site
Write-Host "Connecting to SharePoint Site..." -ForegroundColor Yellow

Connect-PnPOnline -Url $SiteURL.Url `
                  -ClientId $env:DigitalSupportAppID `
                  -Tenant 'mhud.onmicrosoft.com' `
                  -Thumbprint $env:DigitalSupportCertificateThumbprint

""
$RawSourceURL = Read-Host "Please paste the folder URL you want to Move"
""
$RawTargetURL = Read-Host "Please paste the folder of the target directory"

# Create a Uri object from the URL
$Sourceuri = [System.Uri]::new($RawSourceURL)
$Targeturi = [System.Uri]::new($RawTargetURL)

# Get the absolute path which will include the percent-encoded spaces
$SourceURL = $Sourceuri.LocalPath
$TargetURL = $Targeturi.LocalPath
""
Write-Host "Source: $SourceURL" -ForegroundColor Green
Write-Host "Target: $TargetURL" -ForegroundColor Green
""
#Get all Items from the Document Library
$Items = Get-PnPFolder -Url $SourceURL | Where-Object {$_.Name -ne "Forms"}
 
#Move All Files and Folders Between Document Libraries
Foreach($Item in $Items)
{
    Move-PnPFile -SourceUrl $Item.ServerRelativeUrl -TargetUrl $TargetURL -AllowSchemaMismatch -Force -AllowSmallerVersionLimitOnDestination -Verbose
    Write-host "Moved Item:"$Item.ServerRelativeUrl
}
