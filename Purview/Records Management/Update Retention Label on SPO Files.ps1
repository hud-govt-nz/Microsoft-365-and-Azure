#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules Microsoft.Graph.Security
#Requires -Modules Microsoft.Graph.Sites
#Requires -Modules Microsoft.Graph.Files

# Update-RetentionLabelsOnSPOFiles.PS1
# A script to update retention labels on files in all SharePoint document libraries based on a CSV mapping file
# The CSV should have two columns: "ExistingLabel" and "NewLabel"
# V2.1 14-Jan-2025

# Run the script with the site url and CSV file path
# .\Update-RetentionLabelsOnSPOFiles.PS1 -SiteUrl "https://<domain>.sharepoint.com/sites/YourSiteName" -CsvPath "C:\Path\To\LabelMappings.csv"

param (
    [Parameter(Mandatory=$true)]
    [string]$siteURL,
    [Parameter(Mandatory=$true)]
    [string]$csvPath
)

function FormatFileSize {
    # Format File Size nicely
    param (
            [parameter(Mandatory = $true)]
            $InFileSize
        ) 
    
    If ($InFileSize -lt 1KB) { # Format the size of a document
        $FileSize = $InFileSize.ToString() + " B" 
    } 
    ElseIf ($InFileSize -lt 1MB) {
        $FileSize = $InFileSize / 1KB
        $FileSize = ("{0:n2}" -f $FileSize) + " KB"
    } 
    Elseif ($InFileSize -lt 1GB) {
        $FileSize = $InFileSize / 1MB
        $FileSize = ("{0:n2}" -f $FileSize) + " MB" 
    }
    Elseif ($InFileSize -ge 1GB) {
        $FileSize = $InFileSize / 1GB
        $FileSize = ("{0:n2}" -f $FileSize) + " GB" 
    }
    Return $FileSize
} 

function Lock-Record {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $driveItemId
    )
    $params = @{ retentionSettings = @{ isRecordLocked = $true }}
    Update-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $driveItemId -BodyParameter $params   
}

# Function to recursively process all items in a library or folder
function Get-DriveItems {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Drive,
        [Parameter()]
        $FolderId,
        [Parameter()]
        $LabelMappings
    )
    # Get data for a folder and its children
    [array]$Data = Get-MgDriveItemChild -DriveId $Drive -DriveItemId $FolderId -All
    
    # Split the data into files and folders
    [array]$Folders = $Data | Where-Object {$_.folder.childcount -gt 0} | Sort-Object Name
    $Global:TotalFolders = $TotalFolders + $Folders.Count
    [array]$Files = $Data | Where-Object {$null -ne $_.file.mimetype} 

    # Process the files
    ForEach ($File in $Files) {  
        # Get retention label information from file
        Try {
            $RetentionLabelName = $null; $RetentionLabelInfo = $null
            $RetentionlabelInfo = Get-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.id
            $RetentionLabelName = $RetentionLabelInfo.Name

            # Check if current label exists in our mapping
            $labelMapping = $LabelMappings | Where-Object { $_.ExistingLabel -eq $RetentionLabelName }
            
            if ($null -ne $RetentionLabelName -and $null -ne $labelMapping) {
                $newRetentionLabelName = $labelMapping.NewLabel
                
                # Check if this is a record label
                $RetentionLabel = $RetentionLabels | Where-Object { $_.DisplayName -eq $RetentionLabelName }
                $isRecordLabel = if ($RetentionLabel.BehaviorDuringRetentionPeriod -eq "retainAsRecord") {$true} else {$false}

                if ($isRecordLabel -and $RetentionLabelInfo.RetentionSettings.IsRecordLocked -ne $true ) {
                    Lock-Record -Drive $Drive -DriveItemId $File.Id
                    $lockedByScript = $true
                } else {
                    $lockedByScript = $false
                }

                if ($newRetentionLabelName -eq "CLEAR") {
                    Remove-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id
                } else {
                    $UpdatedLabel = Update-MgDriveItemRetentionLabel -DriveId $Drive -DriveItemId $File.Id -BodyParameter @{name = $newRetentionLabelName}
                }  

                If ($File.LastModifiedDateTime) {
                    $LastModifiedDateTime = Get-Date $File.LastModifiedDateTime -format 'dd-MMM-yyyy HH:mm'
                } Else {
                    $LastModifiedDateTime = $null
                }
                If ($File.CreatedDateTime) {
                    $FileCreatedDateTime = Get-Date $File.CreatedDateTime -format 'dd-MMM-yyyy HH:mm'
                }
        
                $ReportLine = [PSCustomObject]@{
                    FileName                = $File.Name
                    Folder                  = $File.parentreference.name
                    Size                    = (FormatFileSize $File.Size)
                    Created                 = $FileCreatedDateTime
                    Author                  = $File.CreatedBy.User.DisplayName
                    LastModified            = $LastModifiedDateTime
                    'Last modified by'      = $File.LastModifiedBy.User.DisplayName
                    'Old Retention label'   = $RetentionLabelName
                    'New Retention label'   = if ($newRetentionLabelName -eq "CLEAR") {""} else {$UpdatedLabel.Name}
                    'Label applied on'      = get-date $UpdatedLabel.labelAppliedDateTime -format 'dd-MMM-yyyy'
                    'Locked by script'      = if ($lockedByScript) {"Yes"} else {"N/A"}
                    WebURL                  = $File.WebUrl
                }
                $ReportData.Add($ReportLine)                            
            }
        } Catch {
            Write-Host ("Error processing file {0}: {1}" -f $File.Name, $_.Exception.Message) 
        }
    }

    # Process the folders
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id -LabelMappings $LabelMappings
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "Sites.Read.All","Files.ReadWrite.All, RecordsManagement.Read.All" -NoWelcome

# Extract hostname and server-relative path from site URL
$uri = [System.Uri]::new($siteURL)
$hostname = $uri.Host
$serverRelativePath = $uri.AbsolutePath.TrimEnd('/')

# Get the site information using the site URL
$site = Get-MgSite -SiteId "$($hostname):$($serverRelativePath)"

if ($null -eq $site) {
    Write-Output "Error: Unable to retrieve site information."
    return
} else {
    $Global:Site = $site
    $siteId = $site.Id
    $SiteName = $site.DisplayName
    Write-Host "Found site to process:" $SiteName 
}

# Get all document libraries (drives) in the site
$drives = Get-MgSiteDrive -SiteId $siteId
if ($null -eq $drives) {
    Write-Output "Error: Unable to retrieve drive information."
    return
} else {
    Write-Host "Found" $drives.Count "drives to process"
}

# Validate and import CSV
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at path: $csvPath"
    return
}

$labelMappings = Import-Csv $csvPath
if (-not ($labelMappings | Get-Member -Name "ExistingLabel") -or -not ($labelMappings | Get-Member -Name "NewLabel")) {
    Write-Error "CSV must contain 'ExistingLabel' and 'NewLabel' columns"
    return
}

Write-Host "Loaded $(($labelMappings | Measure-Object).Count) label mappings from CSV"

# Retrieve retention label from Purview to know if the old label is a record label (retainAsRecord property)
$RetentionLabels = Get-MgSecurityLabelRetentionLabel

[datetime]$StartProcessing = Get-Date
$Global:TotalFolders = 1

# Create output list and CSV file
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Files {0}-AllLibraries.csv" -f $Site.displayName)

# Process each drive in the site
foreach ($drive in $drives) {
    $DriveName = $drive.Name
    Write-Host "`nProcessing drive:" $DriveName
    Write-Host "Fetching file information..."
    
    # Get the items in the root, including child folders
    Get-DriveItems -Drive $Drive.Id -FolderId "root" -LabelMappings $labelMappings
}

[datetime]$EndProcessing = Get-Date
$ElapsedTime = ($EndProcessing - $StartProcessing)
$FilesPerMinute = [math]::Round(($ReportData.Count / ($ElapsedTime.TotalSeconds / 60)), 2)
Write-Host ""
Write-Host ("Processed {0} files in {1} folders across {2} libraries in {3}:{4} minutes ({5} files/minute)" -f `
   $ReportData.Count, $TotalFolders, $drives.Count, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)

Write-Host ""
Write-Host "Retention Labels updated"
$ReportData | Group-Object 'Old Retention label' -NoElement | Sort-Object Count -Descending | Format-Table Name, Count
$ReportData | Out-GridView -Title ("Updated retention labels on files in all document libraries for the {0} site" -f $SiteName)
$ReportData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("Report data saved to {0}" -f $CSVOutputFile)

# Disconnect from Microsoft Graph
Disconnect-MgGraph