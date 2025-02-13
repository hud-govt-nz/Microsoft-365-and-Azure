#=============================================================================
# Script Name: SPO_Labels_Report.ps1
# Created: 10.02.2025
# Author: Ashley Forde
# 
# Description:
#   This script generates a comprehensive report of SharePoint Online files that have
#   retention or sensitivity labels applied. It provides flexible scanning options
#   including filtering by specific compliance tags and selecting target sites.
#
# Features:
#   - Filter by specific compliance tag(s) or scan all tagged files
#   - Interactive site selection via GUI
#   - Real-time XLSX export for immediate results
#   - Progress tracking for long operations
#   - Detailed file information including library and folder paths
#
# Parameters:
#   -ComplianceTagScope : String
#       Optional. Specify "all" to scan all tagged files, or provide a comma-separated
#       list of specific tag names to filter by. Default: "all"
#   
#   -SelectSites : Switch
#       Optional. When enabled, shows a GUI dialog to select specific sites to scan.
#       Default: False (scans all sites)
#
# Example Usage:
#   # Scan all sites for any tagged files
#   .\SPO_Labels_Report.ps1
#
#   # Scan selected sites for specific compliance tags
#   .\SPO_Labels_Report.ps1 -ComplianceTagScope "Confidential,Internal" -SelectSites
#=============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComplianceTagScope = @("all"),
    
    [Parameter(Mandatory = $false)]
    [switch]$SelectSites = $false
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
$fileName = "labelsReport" + $dateTime
$outputPath = "$directoryPath\Reports" + "\" + $fileName + ".xlsx"
$transcriptPath = "$directoryPath\Logs" + "\" + $fileName + "_transcript.txt"

# Import the ImportExcel module
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

# Function to test file size and create a new file if the current file exceeds 10MB
function Test-FileSize {
    param (
        [string]$FilePath,
        [string]$BaseFileName,
        [string]$DirectoryPath,
        [int]$MaxFileSizeMB = 10
    )
    
    $fileInfo = Get-Item $FilePath -ErrorAction SilentlyContinue
    if ($fileInfo -and $fileInfo.Length -gt ($MaxFileSizeMB * 1MB)) {
        $dateTime = (Get-Date).ToString("dd-MM-yyyy-hh-ss")
        $newFileName = "$BaseFileName-$dateTime.xlsx"
        $newFilePath = Join-Path $DirectoryPath "Reports\$newFileName"
        return $newFilePath
    }
    return $FilePath
}

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

# Define libraries to exclude from scanning
$ExcludedLibraries = @(
    "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Pages", "Settings", "Videos",
    "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office"
)

#Function to convert bytes to GB
function Convert-ToGB {
    param([double]$bytes)
    return [math]::Round(($bytes / 1GB), 2)
}
   
# Main function to report on labeled files within a site
function ReportFileLabels($siteUrl) {
    # Disable PnP PowerShell update check
    $env:PNPPOWERSHELL_UPDATECHECK = "Off"
    # Connect to the specific site
    Write-Host "  Connecting to site: $siteUrl" -ForegroundColor Gray
    Connect-PnPOnline -url $siteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $siteconn = Get-PnPConnection
    
    try {
        # Get all document libraries, excluding system libraries
        Write-Host "  Retrieving document libraries..." -ForegroundColor Gray
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
        }

        # Process each library
        foreach ($library in $DocLibraries) {
            Write-Host "    Scanning library: $($library.Title)" -ForegroundColor Yellow
            
            # Retrieve all items with additional fields needed for the report
            Get-PnPListItem -List $library.Title -Fields "ID","GUID","ParentUniqueId","_UIVersionString","_ComplianceTag","_DisplayName","FileDirRef","FileLeafRef","FileRef","Author","Editor","Created","Last_x0020_Modified","File_x0020_Size" -PageSize 1000 -Connection $siteconn | ForEach-Object {
                # Filter items based on label presence and compliance tag scope
                if (($_.FieldValues["_DisplayName"] -or $_.FieldValues["_ComplianceTag"]) -and
                    ($ComplianceTags.Count -eq 0 -or $_.FieldValues["_ComplianceTag"] -in $ComplianceTags)) {

                    # Get file size values (assuming File_x0020_Size holds file size in bytes)
                    $FileSizeBytes = $_.FieldValues["File_x0020_Size"]
                    # Use the same value for TotalSizeBytes or adjust as needed
                    $TotalSizeBytes = $FileSizeBytes
                    
                    # Create item object with relevant information
                    $item = [PSCustomObject]@{
                        SiteUrl             = $siteUrl
                        Library             = $library.Title
                        FolderPath          = $_.FieldValues["FileDirRef"]
                        Title               = $_.FieldValues["FileLeafRef"]
                        ID                  = $_.Id
                        ServerRelativePath  = $_.FieldValues["FileRef"]
                        RetentionLabel      = $_.FieldValues["_ComplianceTag"]
                        SensitivityLabel    = $_.FieldValues["_DisplayName"]
                        Created             = $_["Created"]
                        CreatedBy           = $_["Author"].LookupValue
                        LastModified        = $_["Last_x0020_Modified"]
                        ModifiedBy         = $_["Editor"].LookupValue
                        FileSizeGB          = Convert-ToGB $FileSizeBytes
                        TotalFileSizeGB     = Convert-ToGB $TotalSizeBytes
                        Version             = $_.FieldValues["_UIVersionString"]
                        UniqueId            = $_.FieldValues["GUID"]
                        ParentFolderUniqueId= $_.FieldValues["ParentUniqueId"]

                    }

                    # Test file size and create a new file if needed
                    $outputPath = Test-FileSize -FilePath $outputPath -BaseFileName $fileName -DirectoryPath $directoryPath

                    # Export item to XLSX in real-time
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
    } catch {
        Write-Output "  Error processing site: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Ensure output directories exist
Initialize-OutputDirectories -ReportPath $outputPath -LogPath $transcriptPath

# Retrieve all SharePoint sites
Write-Host "Retrieving SharePoint sites..." -ForegroundColor Cyan
$allSites = Get-PnPTenantSite -Filter "Url -like '$TenantURL'" -Connection $adminConnection |
    Where-Object { $_.Template -ne 'RedirectSite#0' }

if ($allSites.Count -eq 0) {
    Write-Output "No sites found in the tenant" -ForegroundColor Yellow
    exit
}

# Handle site selection based on parameter
Write-Host "Processing site selection..." -ForegroundColor Cyan
$sites = if ($SelectSites) {
    Write-Host "Opening site selection dialog..." -ForegroundColor Yellow
    $selectedSites = $allSites | Select-Object Title, Url | Out-GridView -Title "Select SharePoint Sites to Process" -OutputMode Multiple
    $allSites | Where-Object { $_.Url -in $selectedSites.Url }
} else {
    Write-Host "Processing all available sites..." -ForegroundColor Yellow
    $allSites
}

if ($sites.Count -eq 0) {
    Write-Output "No sites selected for processing" -ForegroundColor Yellow
    exit
}

# Process selected sites with progress tracking
Write-Host "`nStarting site processing..." -ForegroundColor Green
$total = $sites.Count
$count = 0

foreach ($site in $sites) {
    $count++
    Write-Progress -Activity "Processing Sites" `
                   -Status "Processing site: $($site.Url) ($count of $total)" `
                   -PercentComplete (($count / $total) * 100)
    Write-Host "Processing Site ($count of $total):" $site.Url -ForegroundColor Magenta
    
    ReportFileLabels -siteUrl $site.Url
}

# Display completion summary
Write-Host "`nReport generation complete!" -ForegroundColor Green
Write-Host "Report location: $outputPath" -ForegroundColor Green
Write-Host "Transcript location: $transcriptPath" -ForegroundColor Green
Write-Host "Total sites processed: $count" -ForegroundColor Green

# Stop transcript logging
Stop-Transcript
