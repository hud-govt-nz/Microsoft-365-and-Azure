#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules Microsoft.Graph.Security
#Requires -Modules Microsoft.Graph.Sites
#Requires -Modules Microsoft.Graph.Files

# Update-RetentionLabelsOnSPOFiles.PS1
# A script to update the retention label on files in a SharePoint document library that uses cmdlets from
# the Microsoft Graph SDK. For iterating thru the files and folders in the library and reporting, this script was modeled after 
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-SPOFilesDocumentLibrary.PS1
# V1.0 14-Jan-2025

# Run the script with the site url to process, the library display name, the old retention label and the new retention label.
# .\Update-RetentionLabelsOnSPOFiles.PS1 -SiteUrl "https://<domain>.sharepoint.com/sites/YourSiteName" -DocumentLibraryName "Documents" -OldRetentionLabelName "Budget" -NewRetentionLabelName "Financial Report"

param (
    [string]$siteURL,
    [string]$documentLibraryName,
    [string]$oldRetentionLabelName,
    [string]$newRetentionLabelName
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
        $FolderId
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

                if ($null -ne $RetentionLabelName -and $RetentionLabelName -eq $oldRetentionLabelName) {
                    if ($global:IsRecordLabel -and $RetentionLabelInfo.RetentionSettings.IsRecordLocked -ne $true ) {
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
                Write-Host ("Error reading retention label data from file {0}" -f $File.Name) 
            }
    }

    # Process the folders
    ForEach ($Folder in $Folders) {
        Write-Host ("Processing folder {0} ({1} files/size {2})" -f $Folder.Name, $Folder.folder.childcount, (FormatFileSize $Folder.Size))
        Get-DriveItems -Drive $Drive -FolderId $Folder.Id
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

# Get the document library (drive) information
$drive = Get-MgSiteDrive -SiteId $siteId | Where-Object { $_.name -eq $documentLibraryName }
if ($null -eq $drive) {
    Write-Output "Error: Unable to retrieve drive information."
    return
} else {
    $DriveName = $drive.Name
    Write-Host "Found drive to process:" $DriveName 
}

# Retrieve retention label from Purview to know if the old label is a record label (retainAsRecord property)
$RetentionLabels = Get-MgSecurityLabelRetentionLabel
$RetentionLabel = $RetentionLabels | Where-Object { $_.DisplayName -eq $oldRetentionLabelName }
$Global:IsRecordLabel = if ($RetentionLabel.BehaviorDuringRetentionPeriod -eq "retainAsRecord") {$true} else {$false}

[datetime]$StartProcessing = Get-Date
$Global:TotalFolders = 1

# Create output list and CSV file
$Global:ReportData = [System.Collections.Generic.List[Object]]::new()
$CSVOutputFile =  ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + ("\Files {0}-{1} library.csv" -f $Site.displayName, $DriveName)

# Get the items in the root, including child folders
Write-Host "Fetching file information..."
Get-DriveItems -Drive $Drive.Id -FolderId "root"

[datetime]$EndProcessing = Get-Date
$ElapsedTime = ($EndProcessing - $StartProcessing)
$FilesPerMinute = [math]::Round(($ReportData.Count / ($ElapsedTime.TotalSeconds / 60)), 2)
Write-Host ""
Write-Host ("Processed {0} files in {1} folders in {2}:{3} minutes ({4} files/minute)" -f `
   $ReportData.Count, $TotalFolders, $ElapsedTime.Minutes, $ElapsedTime.Seconds, $FilesPerMinute)

Write-Host ""
Write-Host "Retention Labels updated"
$ReportData | Group-Object 'Old Retention label' -NoElement | Sort-Object Count -Descending | Format-Table Name, Count
$ReportData | Out-GridView -Title ("Updated retention labels on files in {0} document library for the {1} site" -f $DriveName, $SiteName)
$ReportData | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding UTF8
Write-Host ("Report data saved to {0}" -f $CSVOutputFile)

# Disconnect from Microsoft Graph
Disconnect-MgGraph