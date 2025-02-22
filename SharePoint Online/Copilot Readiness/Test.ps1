[CmdletBinding()]
param(
    [Parameter(Mandatory=$false,
    HelpMessage="Domain name for SharePoint tenant (e.g., 'mhud' for mhud.sharepoint.com)")]
    [string]$domain = "mhud",
    
    [Parameter(Mandatory=$false,
    HelpMessage="Show site selection grid to choose specific sites")]
    [switch]$SelectSites
)

$adminSiteURL = "https://$domain-Admin.SharePoint.com"
$TenantURL = "https://$domain.sharepoint.com"
$dateTime = (Get-Date).ToString("dd-MM-yyyy-hh-ss")
$directoryPath = "C:\HUD\06_Reporting\SPO\Test\"
$fileName = "labelsReport" + $dateTime
$logFile = "$directoryPath\scan_log_$dateTime.log"
$rowLimit = 1000000  # Row limit per file part
$currentLabeledPart = 1
$currentUnlabeledPart = 1
$labeledRowCount = 0
$unlabeledRowCount = 0
$sitesBatchSize = 300  # Process 300 sites before writing to CSV
$currentFileSites = @() # Array to store site batch data
$currentSite = 0

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

function GetOutputFilePath([string]$type, [int]$part) {
    return "$directoryPath\${fileName}_${type}_pt$part.csv"
}

function ExportToCSV($items, [bool]$append = $false) {
    # Create directory if it doesn't exist
    if (-not (Test-Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }

    # Split items into labeled and unlabeled
    $labeledItems = $items | Where-Object { -not [string]::IsNullOrEmpty($_.RetentionLabel) }
    $unlabeledItems = $items | Where-Object { [string]::IsNullOrEmpty($_.RetentionLabel) }

    # Handle labeled items
    if ($labeledItems) {
        $labeledOutputFile = GetOutputFilePath "labeled" $currentLabeledPart
        
        if ($append -and (Test-Path $labeledOutputFile)) {
            $existingRows = (Import-Csv $labeledOutputFile | Measure-Object).Count
            $labeledRowCount = $existingRows
        }
        
        foreach ($item in $labeledItems) {
            if ($labeledRowCount -ge $rowLimit) {
                $currentLabeledPart++
                $labeledRowCount = 0
                $labeledOutputFile = GetOutputFilePath "labeled" $currentLabeledPart
                $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation
            } else {
                $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation -Append:($labeledRowCount -gt 0)
            }
            $labeledRowCount++
        }
        Write-Log "Processed labeled items to part $currentLabeledPart (Row count: $labeledRowCount)"
    }

    # Handle unlabeled items
    if ($unlabeledItems) {
        $unlabeledOutputFile = GetOutputFilePath "unlabeled" $currentUnlabeledPart
        
        if ($append -and (Test-Path $unlabeledOutputFile)) {
            $existingRows = (Import-Csv $unlabeledOutputFile | Measure-Object).Count
            $unlabeledRowCount = $existingRows
        }
        
        foreach ($item in $unlabeledItems) {
            if ($unlabeledRowCount -ge $rowLimit) {
                $currentUnlabeledPart++
                $unlabeledRowCount = 0
                $unlabeledOutputFile = GetOutputFilePath "unlabeled" $currentUnlabeledPart
                $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation
            } else {
                $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation -Append:($unlabeledRowCount -gt 0)
            }
            $unlabeledRowCount++
        }
        Write-Log "Processed unlabeled items to part $currentUnlabeledPart (Row count: $unlabeledRowCount)"
    }
}

$allSites | foreach-object {   
    $currentSite++
    Write-Progress -Id 1 -Activity "Processing Sites" -Status "Site $currentSite of $totalSites" -PercentComplete (($currentSite / $totalSites) * 100)
    
    # Collect items from this site
    $siteData = ReportFileLabels -siteUrl $_.Url
    $currentFileSites += $siteData

    # Export to CSV after every 100 sites or on the last site
    if ($currentSite % $sitesBatchSize -eq 0 -or $currentSite -eq $totalSites) {
        Write-Host "`nExporting batch $([math]::Ceiling($currentSite / $sitesBatchSize)) of $([math]::Ceiling($totalSites / $sitesBatchSize))" -ForegroundColor Yellow
        Write-Log "Exporting batch (Sites $($currentSite - $currentFileSites.Count + 1) to $currentSite)" -Level "INFO"
        
        # Determine if we should append or create new file
        $shouldAppend = $currentSite -gt $sitesBatchSize
        ExportToCSV -items $currentFileSites -append $shouldAppend
        
        # Clear the current batch after export
        $currentFileSites = @()
    }
}

Write-Progress -Id 1 -Activity "Processing Sites" -Completed
Write-Host "`nReport generation completed!" -ForegroundColor Green
Write-Log "Report generation completed!" -Level "SUCCESS"
