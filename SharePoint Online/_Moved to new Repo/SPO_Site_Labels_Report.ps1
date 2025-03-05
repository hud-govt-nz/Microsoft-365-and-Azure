#=============================================================================
# Script Name: SPO_Site_Labels_Report.ps1
# Created: 2024
# Author: Ashley Forde
# 
# Description:
#   This script generates a report of SharePoint Online files that have retention 
#   or sensitivity labels applied for a specific site. The site can be selected
#   via parameter or through a GUI selection.
#
# Parameters:
#   -ComplianceTagScope : String
#       Optional. Specify "all" to scan all tagged files, or provide a comma-separated
#       list of specific tag names to filter by. Default: "all"
#   
#   -SiteUrl : String
#       Optional. The specific SharePoint site URL to scan. If not provided, 
#       a selection GUI will appear.
#
# Example Usage:
#   # Select site via GUI
#   .\SPO_Site_Labels_Report.ps1
#
#   # Scan specific site URL with specific compliance tags
#   .\SPO_Site_Labels_Report.ps1 -SiteUrl "https://mhud.sharepoint.com/sites/MyTeam" -ComplianceTagScope "Confidential,Internal"
#=============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComplianceTagScope = @("all"),
    
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl
)

# Clear screen for better visibility
Clear-Host

# Initialize basic configuration
Write-Host "Initializing script configuration..." -ForegroundColor Cyan
$domain = "mhud"
$adminSiteURL = "https://$domain-Admin.SharePoint.com"
$TenantURL = "https://$domain.sharepoint.com"
$dateTime = (Get-Date).ToString("dd-MM-yyyy-hh-ss")
$directoryPath = "C:\HUD\06_Reporting\SPO"
$fileName = "siteLabelReport" + $dateTime
$outputPath = "$directoryPath\Reports" + "\" + $fileName + ".xlsx"
$transcriptPath = "$directoryPath\Logs" + "\" + $fileName + "_transcript.txt"

# Import required module
Import-Module ImportExcel

# Function to ensure output directories exist
function Initialize-OutputDirectories {
    param (
        [string]$ReportPath,
        [string]$LogPath
    )
    
    try {
        $ReportDir = Split-Path $ReportPath -Parent
        $LogDir = Split-Path $LogPath -Parent
        
        if (-not (Test-Path $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
        if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    }
    catch {
        Write-Error "Failed to create output directories: $_"
        throw
    }
}

#Function to convert bytes to GB
function Convert-ToGB {
    param([double]$bytes)
    return [math]::Round(($bytes / 1GB), 2)
}

# Define libraries to exclude from scanning
$ExcludedLibraries = @(
    "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Pages", "Settings", "Videos",
    "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office"
)

# Start transcript logging
Start-Transcript -Path $transcriptPath
Write-Host "Script started at: $(Get-Date)" -ForegroundColor Green
Write-Host "Transcript being saved to: $transcriptPath" -ForegroundColor Green

# Process compliance tag parameter
Write-Host "Setting up compliance tag filtering..." -ForegroundColor Cyan
$ComplianceTags = @()
if ($ComplianceTagScope.Count -eq 1 -and $ComplianceTagScope[0] -eq "all") {
    Write-Host "Scanning for all compliance tags" -ForegroundColor Yellow
} else {
    $ComplianceTags = $ComplianceTagScope
    Write-Host "Filtering for specific tags:" -ForegroundColor Yellow
    $ComplianceTags | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

# Disable PnP PowerShell update check
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

# Connect to SharePoint Online Admin center
Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Cyan
Connect-PnPOnline -Url $adminSiteURL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
$adminConnection = Get-PnPConnection

# If no SiteUrl provided, show selection dialog
if (-not $SiteUrl) {
    Write-Host "No site URL provided, opening site selection dialog..." -ForegroundColor Yellow
    $allSites = Get-PnPTenantSite -Filter "Url -like '$TenantURL'" -Connection $adminConnection |
        Where-Object { $_.Template -ne 'RedirectSite#0' }
    
    $selectedSite = $allSites | 
        Select-Object Title, Url | 
        Out-GridView -Title "Select SharePoint Site to Process" -OutputMode Single
    
    if (-not $selectedSite) {
        Write-Host "No site selected. Exiting..." -ForegroundColor Yellow
        Stop-Transcript
        exit
    }
    
    $SiteUrl = $selectedSite.Url
}

Write-Host "Processing site: $SiteUrl" -ForegroundColor Cyan

# Connect to the specific site
Connect-PnPOnline -url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
$siteconn = Get-PnPConnection

try {
    # Get all document libraries, excluding system libraries
    Write-Host "Retrieving document libraries..." -ForegroundColor Gray
    $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
        $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
    }

    # Initialize progress tracking for libraries
    $libraryCount = 0
    $totalLibraries = $DocLibraries.Count

    # Process each library
    foreach ($library in $DocLibraries) {
        $libraryCount++
        Write-Progress -Activity "Processing Libraries" -Status "Scanning: $($library.Title)" -PercentComplete (($libraryCount / $totalLibraries) * 100)
        Write-Host "Scanning library ($libraryCount of $totalLibraries): $($library.Title)" -ForegroundColor Yellow

        # Get items with labels
        Get-PnPListItem -List $library.Title -Fields "ID","GUID","ParentUniqueId","_UIVersionString","_ComplianceTag","_DisplayName","FileDirRef","FileLeafRef","FileRef","Author","Editor","Created","Last_x0020_Modified","File_x0020_Size" -PageSize 1000 -Connection $siteconn | ForEach-Object {
            if (($_.FieldValues["_DisplayName"] -or $_.FieldValues["_ComplianceTag"]) -and
                ($ComplianceTags.Count -eq 0 -or $_.FieldValues["_ComplianceTag"] -in $ComplianceTags)) {

                $FileSizeBytes = $_.FieldValues["File_x0020_Size"]
                
                # Create and export item
                $item = [PSCustomObject]@{
                    SiteUrl             = $SiteUrl
                    Library            = $library.Title
                    FolderPath         = $_.FieldValues["FileDirRef"]
                    Title              = $_.FieldValues["FileLeafRef"]
                    ID                 = $_.Id
                    ServerRelativePath = $_.FieldValues["FileRef"]
                    RetentionLabel     = $_.FieldValues["_ComplianceTag"]
                    SensitivityLabel   = $_.FieldValues["_DisplayName"]
                    Created            = $_["Created"]
                    CreatedBy          = $_["Author"].LookupValue
                    LastModified       = $_["Last_x0020_Modified"]
                    ModifiedBy         = $_["Editor"].LookupValue
                    FileSizeGB         = Convert-ToGB $FileSizeBytes
                    Version            = $_.FieldValues["_UIVersionString"]
                    UniqueId           = $_.FieldValues["GUID"]
                }

                # Export to Excel
                $ExcelParams = @{
                    Path = $outputPath
                    WorksheetName = "LabeledFiles"
                    AutoSize = $true
                    AutoFilter = $true
                    FreezeTopRow = $true
                    BoldTopRow = $true
                }

                if (Test-Path -Path $outputPath) {
                    $ExcelParams.Add("Append", $true)
                }

                $item | Export-Excel @ExcelParams
            }
        }
    }

    # Display completion summary
    Write-Host "`nReport generation complete!" -ForegroundColor Green
    Write-Host "Report location: $outputPath" -ForegroundColor Green
    Write-Host "Transcript location: $transcriptPath" -ForegroundColor Green
    Write-Host "Libraries processed: $libraryCount" -ForegroundColor Green

} catch {
    Write-Error "Error processing site: $($_.Exception.Message)"
}

# Stop transcript logging
Stop-Transcript