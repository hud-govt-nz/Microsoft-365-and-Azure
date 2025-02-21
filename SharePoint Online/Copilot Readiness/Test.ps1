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
$outputFileLabeled = "$directoryPath\${fileName}_labeled.csv"
$outputFileUnlabeled = "$directoryPath\${fileName}_unlabeled.csv"
$sitesBatchSize = 100  # Process 100 sites before writing to CSV
$currentFileSites = @() # Array to store site batch data

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
    $env:PNPPOWERSHELL_UPDATECHECK = "Off"
    Connect-PnPOnline -url $siteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
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
            
            $items = Get-PnPListItem -List $library.Title -Fields "ID","_ComplianceTag","_DisplayName","FileLeafRef","FileRef","FileDirRef","Last_x0020_Modified","Created_x0020_Date","_UIVersionString","SMTotalFileStreamSize" -PageSize 1000 -Connection $siteconn

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
        Write-Progress -Id 2 -Activity "Processing Libraries" -Completed
        Write-Host "Completed site: $siteUrl - Total documents: $totalSiteDocuments" -ForegroundColor Cyan
        Write-Log "Site completed: $siteUrl - Total documents: $totalSiteDocuments"
        return $siteItems
    } 
    catch {
        Write-Host "Error processing site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error processing site $siteUrl : $($_.Exception.Message)" "ERROR"
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

    if ($append) {
        if ($labeledItems) {
            $labeledItems | Export-Csv -Path $outputFileLabeled -NoTypeInformation -Append
            Write-Log "Appended labeled items to: $outputFileLabeled"
        }
        if ($unlabeledItems) {
            $unlabeledItems | Export-Csv -Path $outputFileUnlabeled -NoTypeInformation -Append
            Write-Log "Appended unlabeled items to: $outputFileUnlabeled"
        }
    } else {
        if ($labeledItems) {
            $labeledItems | Export-Csv -Path $outputFileLabeled -NoTypeInformation
            Write-Log "Created new labeled CSV file: $outputFileLabeled"
        }
        if ($unlabeledItems) {
            $unlabeledItems | Export-Csv -Path $outputFileUnlabeled -NoTypeInformation
            Write-Log "Created new unlabeled CSV file: $outputFileUnlabeled"
        }
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
