[CmdletBinding()]
param(
    [Parameter(Mandatory=$false,
    HelpMessage="Domain name for SharePoint tenant (e.g., 'mhud' for mhud.sharepoint.com)")]
    [string]$domain = "mhud",
    
    [Parameter(Mandatory=$false,
    HelpMessage="Show site selection grid to choose specific sites")]
    [switch]$SelectSites
)

function GetOutputFilePath([string]$type, [int]$part) {
    return "$directoryPath\${fileName}_${type}_pt$part.csv"
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [switch]$Console
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    
    # Always write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Only write to console if Console switch is used
    if ($Console) {
        Write-Host $Message -ForegroundColor $(
            switch ($Level) {
                "INFO" { "White" }
                "WARNING" { "Yellow" }
                "ERROR" { "Red" }
                "SUCCESS" { "Green" }
                "PROGRESS" { "Cyan" }
                default { "White" }
            }
        )
    }
}

function ReportFileLabels($siteUrl) {
    Write-Log "Starting scan of site: $siteUrl"
    Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan
    
    try {
        $env:PNPPOWERSHELL_UPDATECHECK = "Off"
        Connect-PnPOnline -url $siteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint -ErrorAction Stop
        $siteconn = Get-PnPConnection
        $siteItems = @()
        $totalSiteDocuments = 0
        
        try {
            $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
                $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
            }

            if (-not $DocLibraries -or $DocLibraries.Count -eq 0) {
                Write-Host "No eligible document libraries found in site: $siteUrl" -ForegroundColor Yellow
                Write-Log "No eligible document libraries found in site: $siteUrl" "WARNING"
                return @()
            }

            Write-Log "Found $($DocLibraries.Count) eligible document libraries in site: $siteUrl"
            $totalLibraries = $DocLibraries.Count
            $currentLibrary = 0

            foreach ($library in $DocLibraries) {
                $currentLibrary++
                $libraryName = $library.Title
                $libraryDocCount = 0
                Write-Progress -Id 2 -Activity "Processing Libraries" -Status "Library $libraryName ($currentLibrary of $totalLibraries)" -PercentComplete (($currentLibrary / $totalLibraries) * 100)
                Write-Host "  Processing library: $libraryName" -ForegroundColor White
                Write-Log "Processing library: $libraryName"
                
                try {
                    # Add extra error checking for null library
                    if ($null -eq $library -or [string]::IsNullOrEmpty($library.Title)) {
                        Write-Host "  Skipping null or invalid library in site $siteUrl" -ForegroundColor Yellow
                        Write-Log "Skipping null or invalid library in site $siteUrl" "WARNING"
                        continue
                    }

                    $items = Get-PnPListItem -List $library.Title -Fields "ID","_ComplianceTag","_DisplayName","FileLeafRef","FileRef","FileDirRef","Last_x0020_Modified","Created_x0020_Date","_UIVersionString","SMTotalFileStreamSize" -PageSize 1000 -Connection $siteconn -ErrorAction Stop

                    if ($items -and $items.Count -gt 0) {
                        $itemCount = $items.Count
                        $currentItem = 0

                        foreach ($_ in $items) {
                            $currentItem++
                            
                            $fileName = $_.FieldValues["FileLeafRef"]
                            $isExcludedFile = $false
                            if ($fileName) {
                                $isExcludedFile = $ExcludedFilePatterns | Where-Object { $fileName -match $_ }
                            }

                            if ($_.FileSystemObjectType -ne "Folder" -and -not $isExcludedFile) {
                                Write-Progress -Id 3 -Activity "Processing Items in $libraryName" -Status "Item $currentItem of $itemCount" -PercentComplete (($currentItem / $itemCount) * 100)
                                $libraryDocCount++
                                
                                $sizeInKB = if ($_.FieldValues["SMTotalFileStreamSize"]) {
                                    [math]::Round($_.FieldValues["SMTotalFileStreamSize"] / 1KB, 2)
                                } else {
                                    0
                                }

                                $item = [PSCustomObject]@{
                                    SiteUrl            = $siteUrl
                                    FolderPath        = $_.FieldValues["FileDirRef"]
                                    ItemID            = $_.FieldValues["ID"]
                                    FileName          = $fileName
                                    RetentionLabel    = $_.FieldValues["_ComplianceTag"]
                                    SensitivityLabel  = $_.FieldValues["_DisplayName"]
                                    Created           = $_.FieldValues["Created_x0020_Date"]
                                    LastModified      = $_.FieldValues["Last_x0020_Modified"]
                                    Version           = $_.FieldValues["_UIVersionString"]
                                    SizeKB            = $sizeInKB
                                    ServerRelativePath = $_.FieldValues["FileRef"]
                                }
                                
                                $siteItems += $item
                            }
                        }
                    }
                    Write-Host "    Processed $libraryDocCount documents in $libraryName" -ForegroundColor Green
                    Write-Log "Library '$libraryName': Processed $libraryDocCount documents"
                    $totalSiteDocuments += $libraryDocCount
                }
                catch {
                    Write-Host "Error processing library $libraryName in site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Error processing library $libraryName in site $siteUrl : $($_.Exception.Message)" "ERROR"
                    # Continue with next library
                    continue
                }
            }
            Write-Progress -Id 2 -Activity "Processing Libraries" -Completed
            Write-Host "Completed site: $siteUrl - Total documents: $totalSiteDocuments" -ForegroundColor Cyan
            Write-Log "Site completed: $siteUrl - Total documents: $totalSiteDocuments"
            return $siteItems
        }
        catch {
            Write-Host "Error accessing libraries in site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error accessing libraries in site $siteUrl : $($_.Exception.Message)" "ERROR"
            return @()
        }
    }
    catch {
        Write-Host "Error connecting to site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error connecting to site $siteUrl : $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function ExportToCSV($items, [bool]$append = $false) {
    # Create directory if it doesn't exist
    if (-not (Test-Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }

    # Split items into labeled and unlabeled
    $labeledItems = $items | Where-Object { -not [string]::IsNullOrEmpty($_.RetentionLabel) }
    $unlabeledItems = $items | Where-Object { [string]::IsNullOrEmpty($_.RetentionLabel) }

    # Handle labeled items with CSV splitting logic
    foreach ($item in $labeledItems) {
         if ($labeledRowCount -ge $rowLimit) {
             $currentLabeledPart++
             $labeledRowCount = 0
             $labeledOutputFile = GetOutputFilePath "labeled" $currentLabeledPart
         }
         if (-not (Test-Path $labeledOutputFile)) {
             $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation
         } else {
             $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation -Append
         }
         $labeledRowCount++
    }
    Write-Log "Exported labeled items"

    # Handle unlabeled items with CSV splitting logic
    foreach ($item in $unlabeledItems) {
         if ($unlabeledRowCount -ge $rowLimit) {
             $currentUnlabeledPart++
             $unlabeledRowCount = 0
             $unlabeledOutputFile = GetOutputFilePath "unlabeled" $currentUnlabeledPart
         }
         if (-not (Test-Path $unlabeledOutputFile)) {
             $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation
         } else {
             $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation -Append
         }
         $unlabeledRowCount++
    }
    Write-Log "Exported unlabeled items"
}

# Script Variables
$adminSiteURL = "https://$domain-Admin.SharePoint.com"
$TenantURL = "https://$domain.sharepoint.com"
$dateTime = (Get-Date).ToString("dd-MM-yyyy-hh-ss")
$directoryPath = "C:\HUD\06_Reporting\SPO\Test\"
$fileName = "labelsReport" + $dateTime
$logFile = "$directoryPath\scan_log_$dateTime.log"
$rowLimit = 900000 # Row limit per file part
$currentLabeledPart = 1
$currentUnlabeledPart = 1
$labeledRowCount = 0
$unlabeledRowCount = 0
$currentSite = 0

# Exclude certain libraries
$ExcludedLibraries = @(
    "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages", "Images", "Pages", "Settings", "Videos",
    "Site Collection Documents", "Site Collection Images", "Style Library", "AppPages", "Apps for SharePoint", "Apps for Office"
)

# Add exclusion patterns for temporary files
$ExcludedFilePatterns = @(
    '~$',        # Temporary Office files
    '.tmp$',     # Temporary files
    '.TMP$',
    '.lck$',     # Lock files
    '.lock$',
    '.part$',    # Partial downloads
    '.crdownload$' # Chrome download temporaries
)

# Initialize CSV output files at the very beginning
$labeledOutputFile = GetOutputFilePath "labeled" 1
$unlabeledOutputFile = GetOutputFilePath "unlabeled" 1
if (Test-Path $labeledOutputFile) { Remove-Item $labeledOutputFile -Force }
if (Test-Path $unlabeledOutputFile) { Remove-Item $unlabeledOutputFile -Force }

Write-Host "Starting SharePoint site scan..." -ForegroundColor Cyan

# Initialize PnP connection
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
Connect-PnPOnline -Url $adminSiteURL -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
$adminConnection = Get-PnPConnection

# Get total number of sites first for progress tracking
$allSites = Get-PnPTenantSite -Filter "Url -like '$TenantURL'" -Connection $adminConnection | Where-Object { $_.Template -ne 'RedirectSite#0' }

# If SelectSites is specified, show site selection grid
if ($SelectSites) {
    Write-Host "Opening site selection grid..." -ForegroundColor Cyan
    $selectedSites = $allSites | Select-Object Url, Title, Template, LastContentModifiedDate | 
        Sort-Object LastContentModifiedDate -Descending |
        Out-GridView -Title "Select Sites to Process (Multiple selection allowed)" -OutputMode Multiple
    if ($selectedSites) {
        $allSites = $allSites | Where-Object { $_.Url -in $selectedSites.Url }
        Write-Host "Selected $($selectedSites.Count) sites to process" -ForegroundColor Green
    } else {
        Write-Host "No sites were selected. Exiting..." -ForegroundColor Yellow
        return
    }
}

$totalSites = $allSites.Count
$currentSite = 0

Write-Host "Found $totalSites sites to process" -ForegroundColor Cyan

$allSites | foreach-object {   
    $currentSite++
    Write-Progress -Id 1 -Activity "Processing Sites" -Status "Site $currentSite of $totalSites" -PercentComplete (($currentSite / $totalSites) * 100)
    
    # Collect items from this site
    $siteData = ReportFileLabels -siteUrl $_.Url
    $currentFileSites += $siteData

    # Immediately export each scanned site by appending to the CSV files as needed
    ExportToCSV -items $siteData -append $true
}

Write-Progress -Id 1 -Activity "Processing Sites" -Completed
Write-Host "`nReport generation completed!" -ForegroundColor Green
Write-Log "Report generation completed!" -Level "SUCCESS"