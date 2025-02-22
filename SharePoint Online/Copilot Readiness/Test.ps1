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
$sitesBatchSize = 2  # Process sites in batches
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

# Add job handling variables
$scanJob = $null
$processingBatch = 1
$totalBatches = [math]::Ceiling($totalSites / $sitesBatchSize)

Write-Host "Found $totalSites sites to process ($totalBatches batches)" -ForegroundColor Cyan

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
        $totalSiteDocuments = 0
        
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
        }

        if (-not $DocLibraries) {
            Write-Log "No eligible document libraries found in site: $siteUrl" "WARNING"
            return
        }

        foreach ($library in $DocLibraries) {
            $libraryName = $library.Title
            $libraryDocCount = 0
            
            if ([string]::IsNullOrEmpty($libraryName)) { continue }

            try {
                $items = Get-PnPListItem -List $libraryName -Fields "ID","_ComplianceTag","_DisplayName","FileLeafRef","FileRef","FileDirRef","Last_x0020_Modified","Created_x0020_Date","_UIVersionString","SMTotalFileStreamSize" -PageSize 1000 -Connection $siteconn
                
                foreach ($item in $items) {
                    if ($item.FileSystemObjectType -eq "Folder") { continue }
                    
                    $fileName = $item.FieldValues["FileLeafRef"]
                    if (-not $fileName -or ($ExcludedFilePatterns | Where-Object { $fileName -match $_ })) { continue }
                    
                    $libraryDocCount++
                    $docItem = [PSCustomObject]@{
                        SiteUrl = $siteUrl
                        FolderPath = $item.FieldValues["FileDirRef"]
                        ItemID = $item.FieldValues["ID"]
                        FileName = $fileName
                        RetentionLabel = $item.FieldValues["_ComplianceTag"]
                        SensitivityLabel = $item.FieldValues["_DisplayName"]
                        Created = $item.FieldValues["Created_x0020_Date"]
                        LastModified = $item.FieldValues["Last_x0020_Modified"]
                        Version = $item.FieldValues["_UIVersionString"]
                        SizeKB = if ($item.FieldValues["SMTotalFileStreamSize"]) { [math]::Round($item.FieldValues["SMTotalFileStreamSize"] / 1KB, 2) } else { 0 }
                        ServerRelativePath = $item.FieldValues["FileRef"]
                    }
                    
                    # Stream the item directly to CSV
                    ExportToCSV -items @($docItem)
                }
                
                Write-Log "Library '$libraryName': Processed $libraryDocCount documents"
                $totalSiteDocuments += $libraryDocCount
            }
            catch {
                Write-Host "Error processing library $libraryName in site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
                Write-Log "Error processing library $libraryName in site $siteUrl : $($_.Exception.Message)" "ERROR"
                continue
            }
        }
        
        Write-Log "Site completed: $siteUrl - Total documents: $totalSiteDocuments"
    }
    catch {
        Write-Host "Error connecting to site $siteUrl : $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error connecting to site $siteUrl : $($_.Exception.Message)" "ERROR"
    }
}

function GetOutputFilePath([string]$type, [int]$part) {
    $filePath = "$directoryPath\${fileName}_${type}_pt$part.csv"
    return $filePath
}

function ShouldCreateNewFile($filePath) {
    if (-not (Test-Path $filePath)) { return $false }
    $file = Get-Item $filePath
    return $file.Length -gt 100MB
}

function ExportToCSV($items) {
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
        
        foreach ($item in $labeledItems) {
            if (ShouldCreateNewFile $labeledOutputFile) {
                $currentLabeledPart++
                $labeledOutputFile = GetOutputFilePath "labeled" $currentLabeledPart
                $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation
            } else {
                $item | Export-Csv -Path $labeledOutputFile -NoTypeInformation -Append
            }
        }
    }

    # Handle unlabeled items
    if ($unlabeledItems) {
        $unlabeledOutputFile = GetOutputFilePath "unlabeled" $currentUnlabeledPart
        
        foreach ($item in $unlabeledItems) {
            if (ShouldCreateNewFile $unlabeledOutputFile) {
                $currentUnlabeledPart++
                $unlabeledOutputFile = GetOutputFilePath "unlabeled" $currentUnlabeledPart
                $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation
            } else {
                $item | Export-Csv -Path $unlabeledOutputFile -NoTypeInformation -Append
            }
        }
    }
}

# Process sites in parallel
$currentSite = 0
$totalSites = $allSites.Count
$maxParallelJobs = 3
$runningJobs = @{}

Write-Host "Found $($allSites.Count) sites to process" -ForegroundColor Cyan
Write-Log "Starting processing of $($allSites.Count) sites" -Level "INFO"

foreach ($site in $allSites) {
    $currentSite++
    Write-Progress -Id 1 -Activity "Processing Sites" -Status "Site $currentSite of $totalSites" -PercentComplete (($currentSite / $totalSites) * 100)
    
    while ($runningJobs.Count -ge $maxParallelJobs) {
        $completedJobs = $runningJobs.Keys | Where-Object { $runningJobs[$_].State -eq 'Completed' }
        foreach ($jobId in $completedJobs) {
            $job = $runningJobs[$jobId]
            Remove-Job -Job $job
            $runningJobs.Remove($jobId)
        }
        if ($runningJobs.Count -ge $maxParallelJobs) {
            Start-Sleep -Seconds 2
        }
    }
    
    ReportFileLabels -siteUrl $site.Url
}

Write-Progress -Id 1 -Activity "Processing Sites" -Completed
Write-Host "`nReport generation completed!" -ForegroundColor Green
Write-Log "Report generation completed!" -Level "SUCCESS"
