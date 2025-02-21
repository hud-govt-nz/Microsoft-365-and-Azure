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
$maxFileSize = 15MB # Target 15MB per file
$sitesPerFile = 15 # Process 15 sites before checking file size
$partNumber = 1
$currentOutputFile = "$directoryPath\$fileName-part$partNumber.xlsx"
$sitesProcessedInCurrentFile = 0

Write-Host "Creating new file every 15 sites or when file reaches 15MB" -ForegroundColor Cyan

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
            Write-Host "File size ($([math]::Round($fileSize / 1MB, 2))MB) exceeded 15MB limit. Creating new file." -ForegroundColor Yellow
            return $true
        }
    }
    return $false
}

function ReportFileLabels($siteUrl) {
    Connect-PnPOnline -url $siteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
    $siteconn = Get-PnPConnection
    try {
        $DocLibraries = Get-PnPList -Includes BaseType, Hidden, Title -Connection $siteconn | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLibraries
        }

        if (-not $DocLibraries -or $DocLibraries.Count -eq 0) {
            Write-Host "No eligible document libraries found in site: $siteUrl" -ForegroundColor Yellow
            return
        }

        $totalLibraries = $DocLibraries.Count
        $currentLibrary = 0
        $totalItemsProcessed = 0

        $DocLibraries | ForEach-Object {
            $currentLibrary++
            $libraryName = $_.Title
            Write-Progress -Id 2 -Activity "Processing Libraries" -Status "Library $libraryName ($currentLibrary of $totalLibraries)" -PercentComplete (($currentLibrary / $totalLibraries) * 100)
            Write-Host "Processing Document Library:" $libraryName "($currentLibrary of $totalLibraries)" -ForegroundColor Yellow
            
            $library = $_
            $items = Get-PnPListItem -List $library.Title -Fields "ID","_ComplianceTag","_DisplayName","FileLeafRef","FileRef","FileDirRef","Last_x0020_Modified","Created_x0020_Date","_UIVersionString","SMTotalFileStreamSize" -PageSize 1000 -Connection $siteconn

            if (-not $items -or $items.Count -eq 0) {
                Write-Host "No items found in library: $libraryName" -ForegroundColor Yellow
                return
            }

            $itemCount = $items.Count
            $currentItem = 0

            $items | ForEach-Object {
                $currentItem++
                $totalItemsProcessed++
                
                # Skip folder items and temporary files
                $fileName = $_.FieldValues["FileLeafRef"]
                $isExcludedFile = $false
                if ($fileName) {
                    $isExcludedFile = $ExcludedFilePatterns | Where-Object { $fileName -match $_ }
                }

                if ($_.FileSystemObjectType -ne "Folder" -and -not $isExcludedFile) {
                    Write-Progress -Id 3 -Activity "Processing Items in $libraryName" -Status "Item $currentItem of $itemCount" -PercentComplete (($currentItem / $itemCount) * 100)
                    
                    # Convert size from bytes to KB
                    $sizeInKB = if ($_.FieldValues["SMTotalFileStreamSize"]) {
                        [math]::Round($_.FieldValues["SMTotalFileStreamSize"] / 1KB, 2)
                    } else {
                        0
                    }

                    $item = [PSCustomObject]@{
                        SiteUrl           = $siteUrl
                        FolderPath       = $_.FieldValues["FileDirRef"]
                        ItemID           = $_.FieldValues["ID"]
                        FileName         = $fileName
                        RetentionLabel    = $_.FieldValues["_ComplianceTag"]
                        SensitivityLabel  = $_.FieldValues["_DisplayName"]
                        Created          = $_.FieldValues["Created_x0020_Date"]
                        LastModified      = $_.FieldValues["Last_x0020_Modified"]
                        Version          = $_.FieldValues["_UIVersionString"]
                        SizeKB           = $sizeInKB
                        ServerRelativePath = $_.FieldValues["FileRef"]
                    }
                    
                    # Create directory if it doesn't exist
                    if (-not (Test-Path $directoryPath)) {
                        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
                    }

                    # Export the item, either creating a new file or appending
                    if (-not (Test-Path $currentOutputFile)) {
                        $item | Export-Excel -Path $currentOutputFile -WorksheetName "Labels Report" -AutoSize -AutoFilter
                    } else {
                        $item | Export-Excel -Path $currentOutputFile -WorksheetName "Labels Report" -Append
                    }

                    # Check file size after append
                    if (CheckFileSize -filePath $currentOutputFile) {
                        $partNumber++
                        $currentOutputFile = "$directoryPath\$fileName-part$partNumber.xlsx"
                    }
                }
            }
            Write-Progress -Id 3 -Activity "Processing Items" -Completed
            Write-Host "Processed $totalItemsProcessed total items so far" -ForegroundColor Cyan
        }
        Write-Progress -Id 2 -Activity "Processing Libraries" -Completed
    } catch {
        Write-Host "An exception was thrown: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Error occurred at line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
}

$allSites | foreach-object {   
    $currentSite++
    Write-Progress -Id 1 -Activity "Processing Sites" -Status "Site $currentSite of $totalSites" -PercentComplete (($currentSite / $totalSites) * 100)
    Write-Host "`nProcessing Site ($currentSite of $totalSites):" $_.Url -ForegroundColor Magenta
    ReportFileLabels -siteUrl $_.Url
    
    $sitesProcessedInCurrentFile++
    if ($sitesProcessedInCurrentFile -ge $sitesPerFile) {
        Write-Host "Processed $sitesPerFile sites. Creating new file." -ForegroundColor Yellow
        $partNumber++
        $currentOutputFile = "$directoryPath\$fileName-part$partNumber.xlsx"
        $sitesProcessedInCurrentFile = 0
    }
}
Write-Progress -Id 1 -Activity "Processing Sites" -Completed

Write-Host "`nReport generation completed!" -ForegroundColor Green
