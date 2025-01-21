Clear-Host
Write-Host '## OneDrive User File Type Report ##' -ForegroundColor Yellow

function Get-UserDriveId {
    param (
        [string]$UserId,
        [string]$token
    )

    $url = "https://graph.microsoft.com/v1.0/users/$UserId/drive"
    try {
        $response = Invoke-RestMethod -Method GET -Uri $url -Headers @{ Authorization = "Bearer $token" }
        return $response.id
    }
    catch {
        Write-Error "Failed to retrieve drive ID for user $($UserId): $_"
        return $null
    }
}
function Get-DriveItems {
    param (
        [string]$DriveId,
        [string]$token,
        [string]$ItemId = "root",
        [ref]$fileCount = 0
    )

    $url = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/children"

    $fileCounts = @{
        "Word"       = 0
        "Excel"      = 0
        "PowerPoint" = 0
        "PDFs"       = 0
        "Text"       = 0
        "Video"      = 0
        "Audio"      = 0
        "OneNote"    = 0
        "SnagIT"     = 0
        "Other"      = 0
    }

    do {
        try {
            $response = Invoke-RestMethod -Method GET -Uri $url -Headers @{ Authorization = "Bearer $token" }
            foreach ($item in $response.value) {
                $fileCount.Value++
                $percentComplete = [math]::Min((($fileCount.Value / 1000) * 100), 100)
                Write-Progress -Activity "Scanning files" -Status "Processing $($fileCount.Value) files" -PercentComplete $percentComplete
                
                if ($item.folder) {
                    # If the item is a folder, recursively call Get-DriveItems
                    $subFolderCounts           = Get-DriveItems -DriveId $DriveId -token $token -ItemId $item.id -fileCount ([ref]$fileCount.Value)
                    $fileCounts["Word"]       += $subFolderCounts["Word"]
                    $fileCounts["Excel"]      += $subFolderCounts["Excel"]
                    $fileCounts["PowerPoint"] += $subFolderCounts["PowerPoint"]
                    $fileCounts["PDF"]       += $subFolderCounts["PDF"]
                    $fileCounts["Text"]       += $subFolderCounts["Text"]
                    $fileCounts["Video"]      += $subFolderCounts["Video"]
                    $fileCounts["Audio"]      += $subFolderCounts["Audio"]
                    $fileCounts["OneNote"]    += $subFolderCounts["OneNote"]
                    $fileCounts["SnagIT"]     += $subFolderCounts["SnagIT"]
                    $fileCounts["Other"]      += $subFolderCounts["Other"]
                } else {
                    # If the item is a file, count it based on its extension
                    switch -Wildcard ($item.name) {
                        "*.DOC" { $fileCounts["Word"]++ }
                        "*.DOCX" { $fileCounts["Word"]++ }
                        "*.RTF" { $fileCounts["Word"]++ }
                        "*.ODT" { $fileCounts["Word"]++ }
                        "*.DOT" { $fileCounts["Word"]++ }
                        "*.DOTX" { $fileCounts["Word"]++ }
                        "*.DOTM" { $fileCounts["Word"]++ }
                        "*.OTT" { $fileCounts["Word"]++ }

                        "*.XLS" { $fileCounts["Excel"]++ }
                        "*.XLSX" { $fileCounts["Excel"]++ }
                        "*.XLSM" { $fileCounts["Excel"]++ }
                        "*.XLTX" { $fileCounts["Excel"]++ }
                        "*.XLTM" { $fileCounts["Excel"]++ }
                        "*.XLSB" { $fileCounts["Excel"]++ }
                        "*.CSV" { $fileCounts["Excel"]++ }

                        "*.PPT" { $fileCounts["Excel"]++ }
                        "*.PPTX" { $fileCounts["Excel"]++ }
                        "*.PPTM" { $fileCounts["Excel"]++ }
                        "*.POTX" { $fileCounts["Excel"]++ }
                        "*.POTM" { $fileCounts["Excel"]++ }
                        "*.PPSX" { $fileCounts["Excel"]++ }
                        "*.PPSM" { $fileCounts["Excel"]++ }

                        "*.PDF" { $fileCounts["PDF"]++ }

                        "*.TXT" { $fileCounts["Text"]++ }
                        "*.LOG" { $fileCounts["Text"]++ }

                        "*.MP4" { $fileCounts["Video"]++ }
                        "*.AVI" { $fileCounts["Video"]++ }
                        "*.MOV" { $fileCounts["Video"]++ }
                        "*.WMV" { $fileCounts["Video"]++ }
                        "*.FLV" { $fileCounts["Video"]++ }
                        "*.MKV" { $fileCounts["Video"]++ }
                        "*.WEBM" { $fileCounts["Video"]++ }
                        "*.MPEG" { $fileCounts["Video"]++ }
                        "*.MPG" { $fileCounts["Video"]++ }

                        "*.MP3" { $fileCounts["Audio"]++ }
                        "*.WAV" { $fileCounts["Audio"]++ }
                        "*.AAC" { $fileCounts["Audio"]++ }
                        "*.M4A" { $fileCounts["Audio"]++ }
                        "*.FLAC" { $fileCounts["Audio"]++ }
                        "*.OGG" { $fileCounts["Audio"]++ }
                        "*.WMA" { $fileCounts["Audio"]++ }
                        "*.MIDI" { $fileCounts["Audio"]++ }
                        "*.AIFF" { $fileCounts["Audio"]++ }

                        "*.ONE" { $fileCounts["OneNote"]++ }
                        "*.ONETOC2" { $fileCounts["OneNote"]++ }
                        "*.ONEPKG" { $fileCounts["OneNote"]++ }
                        "*.NOTE" { $fileCounts["OneNote"]++ }

                        "*.SNAG" { $fileCounts["SnagIT"]++ }
                        "*.SNAGX" { $fileCounts["SnagIT"]++ }
                        "*.SNAGPROJ" { $fileCounts["SnagIT"]++ }
                        
                        default  { $fileCounts["Other"]++ }
                    }
                }
            }
            $url = $response.'@odata.nextLink'
        }
        catch {
            Write-Error "Failed to retrieve items for drive $($DriveId): $_"
            break
        }
    } while ($url)

    return $fileCounts
}

# Connect to Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token = $CollectToken.RequestMessage.Headers.Authorization.Parameter

        
    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

# Collect all licensed users
$AllUsers = Get-MgBetaUser -All | Where-Object { $_.AssignedLicenses.SkuId -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06"} | Select-Object DisplayName, UserPrincipalName

# Initialize results table
$results = @()

# Initialize user progress  
$totalUsers  = $AllUsers.Count
$currentUser = 0

# Select specific user(s) for testing purposes
$selectedUsersInput = Read-Host "Enter the user principal names (comma-separated)"
$selectedUsers = $selectedUsersInput -split ',' | ForEach-Object { $_.Trim() }

# Filter users based on the selected user principal names
$AllUsers = $AllUsers | Where-Object { $selectedUsers -contains $_.UserPrincipalName }

# Check if there are any users to process
if ($AllUsers.Count -eq 0) {
    Write-Host "No users found with the specified criteria."
    exit
}

# Write output naming first 5 users
$AllUsers | Select-Object DisplayName, UserPrincipalName | Format-Table -AutoSize

foreach ($user in $AllUsers) {  
    $currentUser++  

    # Calculate percent complete and ensure it is within the range of 0 to 100  
    $percentComplete = [math]::Min((($currentUser / $totalUsers) * 100), 100)

    Write-Progress -Activity "Processing users" -Status "Processing user $currentUser of $totalUsers" -PercentComplete $percentComplete  

    $driveId = Get-UserDriveId -UserId $user.UserPrincipalName -token $token
    if ($driveId) {  
        $fileCount  = 0
        $fileCounts = Get-DriveItems -DriveId $driveId -token $token -fileCount ([ref]$fileCount)

        # Initialize file progress
        $totalFiles  = $fileCount
        $currentFile = 0

        foreach ($fileType in $fileCounts.Keys) {
            $currentFile         += $fileCounts[$fileType]
            $filePercentComplete  = [math]::Min((($currentFile / $totalFiles) * 100), 100)

            Write-Progress -Activity "Processing users" -Status "Processing files for user $currentUser of $totalUsers" -PercentComplete $percentComplete -CurrentOperation "Processing $currentFile of $totalFiles files"
        }

        $results += [PSCustomObject]@{  
            UserPrincipalName = $user.UserPrincipalName
            DisplayName       = $user.DisplayName
            WordDocuments     = $fileCounts["Word"]
            ExcelSpreadsheets = $fileCounts["Excel"]
            PowerPointSlides  = $fileCounts["PowerPoint"]
            PDFs              = $fileCounts["PDF"]
            TextFiles         = $fileCounts["Text"]
            VideoFiles        = $fileCounts["Video"]
            AudioFiles        = $fileCounts["Audio"]
            OneNoteFiles      = $fileCounts["OneNote"]
            SnagITFiles       = $fileCounts["SnagIT"]
            OtherFiles        = $fileCounts["Other"]
            TotalFiles        = $fileCount
        }
        }  
    }  


# Clear progress bar
Write-Progress -Activity "Processing users" -Completed

# Output results
$results | Format-Table -Property UserPrincipalName, DisplayName, WordDocuments, ExcelSpreadsheets, PowerPointSlides, PDFs, TextFiles, VideoFiles, AudioFiles, OneNoteFiles, SnagITFiles, OtherFiles, TotalFiles -AutoSize -Wrap

# Disconnect from Graph
Disconnect-MgGraph | Out-Null