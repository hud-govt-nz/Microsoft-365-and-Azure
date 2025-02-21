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
$maxFileSize = 30MB # Target 30MB per file
$sitesPerFile = 50 # Process 50 sites before creating new file
$partNumber = 1
$currentOutputFile = "$directoryPath\$fileName-part$partNumber.csv"
$sitesProcessedInCurrentFile = 0
$currentFileSites = @() # Array to store site data before writing to file

Write-Host "Processing 50 sites per file, creating new file at 30MB" -ForegroundColor Cyan

# Check if ImportExcel module is installed, if not install it
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Initialize PnP connection
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

function CheckFileSize {
    param (
        [string]$filePath
    )
    if (Test-Path $filePath) {
        $fileSize = (Get-Item $filePath).Length
        if ($fileSize -ge $maxFileSize) {
            Write-Host "File size ($([math]::Round($fileSize / 1MB, 2))MB) exceeded 30MB limit. Creating new file." -ForegroundColor Yellow
            return $true
        }
    }
    return $false
}

function ReportFileLabels($siteUrl) {
    Connect-PnPOnline -url $siteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $siteconn = Get-PnPConnection
    $siteItems = @() # Array to store items for this site
    try {
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
        }

        if (-not $DocLibraries -or $DocLibraries.Count -eq 0) {
            Write-Host "No eligible document libraries found in site: $siteUrl" -ForegroundColor Yellow
            return @()
        }

        $totalLibraries = $DocLibraries.Count
        $currentLibrary = 0
        $totalItemsProcessed = 0

        foreach ($library in $DocLibraries) {
            $currentLibrary++
            $libraryName = $library.Title
            Write-Progress -Id 2 -Activity "Processing Libraries" -Status "Library $libraryName ($currentLibrary of $totalLibraries)" -PercentComplete (($currentLibrary / $totalLibraries) * 100)
            Write-Host "Processing Document Library:" $libraryName "($currentLibrary of $totalLibraries)" -ForegroundColor Yellow
            
            $items = Get-PnPListItem -List $library.Title -Fields "ID","_ComplianceTag","_DisplayName","FileLeafRef","FileRef","FileDirRef","Last_x0020_Modified","Created_x0020_Date","_UIVersionString","SMTotalFileStreamSize" -PageSize 1000 -Connection $siteconn

            if ($items -and $items.Count -gt 0) {
                $itemCount = $items.Count
                $currentItem = 0

                foreach ($_ in $items) {
                    $currentItem++
                    $totalItemsProcessed++
                    
                    $fileName = $_.FieldValues["FileLeafRef"]
                    $isExcludedFile = $false
                    if ($fileName) {
                        $isExcludedFile = $ExcludedFilePatterns | Where-Object { $fileName -match $_ }
                    }

                    if ($_.FileSystemObjectType -ne "Folder" -and -not $isExcludedFile) {
                        Write-Progress -Id 3 -Activity "Processing Items in $libraryName" -Status "Item $currentItem of $itemCount" -PercentComplete (($currentItem / $itemCount) * 100)
                        
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
        }
        Write-Progress -Id 2 -Activity "Processing Libraries" -Completed
        Write-Host "Processed $totalItemsProcessed total items" -ForegroundColor Cyan
        return $siteItems
    } catch {
        Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function ExportToExcel($items) {
    # Create directory if it doesn't exist
    if (-not (Test-Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }

    $items | Export-Csv -Path $currentOutputFile -NoTypeInformation

    # Check if file size exceeds limit
    if ((Get-Item $currentOutputFile).Length -ge $maxFileSize) {
        Write-Host "File size exceeded 30MB limit. Creating new file for next batch." -ForegroundColor Yellow
        $script:partNumber++
        $script:currentOutputFile = "$directoryPath\$fileName-part$partNumber.csv"
    }
}

$allSites | foreach-object {   
    $currentSite++
    Write-Progress -Id 1 -Activity "Processing Sites" -Status "Site $currentSite of $totalSites" -PercentComplete (($currentSite / $totalSites) * 100)
    Write-Host "`nProcessing Site ($currentSite of $totalSites):" $_.Url -ForegroundColor Magenta
    
    # Collect items from this site
    $siteData = ReportFileLabels -siteUrl $_.Url
    $currentFileSites += $siteData
    
    $sitesProcessedInCurrentFile++
    if ($sitesProcessedInCurrentFile -ge $sitesPerFile) {
        Write-Host "Processed $sitesPerFile sites. Writing to file." -ForegroundColor Yellow
        ExportToExcel -items $currentFileSites
        $currentFileSites = @() # Clear the array for next batch
        $sitesProcessedInCurrentFile = 0
    }
}

# Export any remaining sites in the last batch
if ($currentFileSites.Count -gt 0) {
    Write-Host "Writing final batch to file." -ForegroundColor Yellow
    ExportToExcel -items $currentFileSites
}

Write-Progress -Id 1 -Activity "Processing Sites" -Completed
Write-Host "`nReport generation completed!" -ForegroundColor Green
