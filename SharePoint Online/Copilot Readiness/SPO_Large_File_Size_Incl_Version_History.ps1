[CmdletBinding()]
param (
    [Parameter()]
    [string]$SiteUrl = "https://mhud-admin.sharepoint.com",
    [Parameter()]
    [string]$directoryPath = "C:\HUD\06_Reporting\SPO",
    [Parameter()]
    [int]$MinFileSize = 100 # Size in MB
)

Clear-Host

# Script Variables
$dateTime = (Get-Date).ToString("yyyy-MM-dd-HHmmss")
$fileName = "SPOLargeFiles_$dateTime"
$outputPath = Join-Path $directoryPath "Reports\$fileName.xlsx"
$transcriptPath = Join-Path $directoryPath "Logs\$fileName.log"
$global:processedSites = 0
$global:totalFiles = 0

# Function to convert bytes to MB
function Convert-ToMB {
    param([double]$bytes)
    return [math]::Round(($bytes / 1MB), 2)
}

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

# Function to check file size and create a new file if the current file exceeds 10MB
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

# Function to process a single file item
function Get-FileData {
    param (
        [Parameter(Mandatory)]
        $Item,
        [Parameter(Mandatory)]
        [string]$SiteUrl,
        [Parameter(Mandatory)]
        [string]$LibraryTitle
    )

    $FileSizeBytes = $Item.FieldValues.SMTotalFileStreamSize
    $TotalSizeBytes = $Item.FieldValues.SMTotalSize.LookupId

    if (($TotalSizeBytes / 1MB) -gt $MinFileSize) {
        $global:totalFiles++
        return [PSCustomObject][ordered]@{
            SiteUrl             = $SiteUrl
            Library            = $LibraryTitle
            FolderPath         = $Item.FieldValues["FileDirRef"]
            Title              = $Item.FieldValues["FileLeafRef"]
            ID                 = $Item.Id
            ServerRelativePath = $Item.FieldValues["FileRef"]
            RetentionLabel     = $Item.FieldValues["_ComplianceTag"]
            SensitivityLabel   = $Item.FieldValues["_DisplayName"]
            Created            = $Item["Created"]
            CreatedBy         = $Item["Author"].LookupValue
            LastModified      = $Item["Last_x0020_Modified"]
            ModifiedBy        = $Item["Editor"].LookupValue
            FileSizeMB        = Convert-ToMB $FileSizeBytes
            TotalFileSizeMB   = Convert-ToMB $TotalSizeBytes
            Version           = $Item.FieldValues["_UIVersionString"]
            UniqueId          = $Item.FieldValues["GUID"]
            ParentFolderUniqueId = $Item.FieldValues["ParentUniqueId"]
        }
    }
}

# Function to process a single library
function Get-DocumentLibrary {
    param (
        [Parameter(Mandatory)]
        $Library,
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )

    Write-Host "Processing library: $($Library.Title)" -ForegroundColor Cyan
    $FileData = @()

    try {
        # Skip processing if library is empty
        if ($Library.ItemCount -eq 0) {
            Write-Host "Skipping empty library: $($Library.Title)" -ForegroundColor Yellow
            return
        }

        $ListItems = Get-PnPListItem -List $Library.Title -Fields FileRef, SMTotalFileStreamSize, SMTotalSize, _UIVersionString, 
            FileLeafRef, Created, Modified, _ComplianceTag, _DisplayName, Author, Editor, GUID, ParentUniqueId -PageSize 2000 `
            -ScriptBlock { 
                Param($items) 
                # Protect against division by zero
                $percentComplete = if ($Library.ItemCount -gt 0) {
                    [math]::Min(($items.Count / $Library.ItemCount * 100), 100)
                } else {
                    100
                }
                Write-Progress -PercentComplete $percentComplete `
                    -Activity "Processing '$($Library.Title)'" `
                    -Status "Processing $($items.Count) of $($Library.ItemCount) items"
            } | Where-Object { $_.FileSystemObjectType -eq "File" }

        foreach ($Item in $ListItems) {
            $FileInfo = Get-FileData -Item $Item -SiteUrl $Site.Url -LibraryTitle $Library.Title
            if ($FileInfo) {
                $FileData += $FileInfo
            }
        }

        # Export batch of files if any found
        if ($FileData.Count -gt 0) {
            # Check file size and create a new file if needed
            $outputPath = Check-FileSize -FilePath $outputPath -BaseFileName $fileName -DirectoryPath $directoryPath

            $ExcelParams = @{
                Path = $outputPath
                WorksheetName = "LargeFiles"
                AutoSize = $true
                AutoFilter = $true
                FreezeTopRow = $true
                BoldTopRow = $true
            }

            if (Test-Path -Path $outputPath) {
                $ExcelParams.Add("Append", $true)
            }

            $FileData | Export-Excel @ExcelParams
        }
    }
    catch {
        Write-Error "Error processing library $($Library.Title): $_"
        # Continue with next library
    }
}

# Main script execution
try {
    Initialize-OutputDirectories -ReportPath $outputPath -LogPath $transcriptPath
    Start-Transcript -Path $transcriptPath

    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow

    # Disable PnP PowerShell update check
    $env:PNPPOWERSHELL_UPDATECHECK = "Off"     
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint


    Write-Host "Retrieving all SharePoint sites..." -ForegroundColor Yellow
    $Sites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like 'https://mhud.sharepoint.com/sites/'"
    
    foreach ($Site in $Sites) {
        $global:processedSites++
        Write-Host "`nProcessing site $global:processedSites of $($Sites.Count): $($Site.Url)" -ForegroundColor Green
        
        try {
            $env:PNPPOWERSHELL_UPDATECHECK = "Off"
            Connect-PnPOnline -Url $Site.Url -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint
            $Libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
            
            foreach ($Library in $Libraries) {
                Get-DocumentLibrary -Library $Library -SiteUrl $Site.Url
            }
        }
        catch {
            Write-Error "Error processing site $($Site.Url): $_"
            # Continue with next site
        }
    }

    # Display completion summary
    Write-Host "`nReport generation complete!" -ForegroundColor Green
    Write-Host "Report location: $outputPath" -ForegroundColor Green
    Write-Host "Transcript location: $transcriptPath" -ForegroundColor Green
    Write-Host "Total sites processed: $global:processedSites" -ForegroundColor Green
    Write-Host "Total large files found: $global:totalFiles" -ForegroundColor Green
}
catch {
    Write-Error "Script execution failed: $_"
}
finally {
    Stop-Transcript
}