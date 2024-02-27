Clear-Host
Write-host "## Microsoft Teams: Users Access Report ##" -ForegroundColor Yellow

# Connect to Microsoft Teams
Connect-MgGraph -NoWelcome | Out-Null
$UPN = (Get-MgContext).Account
Connect-MicrosoftTeams -AccountId $UPN -Confirm:$false | Out-Null

$UserEmail = Read-host  "Please enter UPN"


$AllData = @()
$AllTeams = Get-Team -User $UserEmail
$TotalCount = $AllTeams.Count
$CurrentCount = 0

# Iterate through each team
foreach ($Team in $AllTeams) {
    $CurrentCount++
    $TeamName = $Team.DisplayName
    $GroupId = $Team.GroupId

    # Update progress bar for team processing
    Write-Progress -Activity "Processing Teams" -Status "Processing team $CurrentCount of $TotalCount $TeamName" -PercentComplete (($CurrentCount / $TotalCount) * 100)

    if ($TeamName -notmatch 'AIP Users' -and $TeamName -notmatch 'Te Pae Kōrero') {
        # Get Team Users
        Get-TeamUser -GroupId $GroupId | ForEach-Object {
            if ($_.User -match $UserEmail) {
                $AllData += [PSCustomObject]@{
                    TeamName = $TeamName
                    Name = $_.Name
                    UPN = $_.User
                    Role = $_.Role
                    Channel = $null
                    ChannelRole = $null
                }
            }
        }

        $AllChannels = Get-TeamChannel -GroupId $GroupId
        $TotalChannels = $AllChannels.Count
        $CurrentChannelCount = 0

        # Iterate through each channel
        foreach ($Channel in $AllChannels) {
            $CurrentChannelCount++
            $channelDispName = $Channel.DisplayName

            # Update progress bar for channel processing
            Write-Progress -Activity "Processing Channels" -Status "Processing channel $CurrentChannelCount of $TotalChannels in team $TeamName" -PercentComplete (($CurrentChannelCount / $TotalChannels) * 100)

            # Get Channel Users
            Get-TeamChannelUser -GroupId $GroupId -DisplayName $channelDispName | ForEach-Object {
                if ($_.User -match $UserEmail) {
                    $AllData += [PSCustomObject]@{
                        TeamName = $TeamName
                        Name = $_.Name
                        UPN = $_.User
                        Role = $null
                        Channel = $channelDispName
                        ChannelRole = $_.Role
                    }
                }
            }
        }
    }
}

# Clear the progress bar
Write-Progress -Activity "Processing Completed" -Completed

# Get Date for File Naming
$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Team Users Access"

[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save as"
$SaveFileDialog.FileName = $FileName

# Show the SaveFileDialog and get the selected file path
$SaveFileResult = $SaveFileDialog.ShowDialog()
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFilePath = $SaveFileDialog.FileName
    # Export to Excel with formatting
    $AllData | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow

    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet = $excelPackage.Workbook.Worksheets["$Date"]

    # Autosize columns if needed
    foreach ($column in 1..$worksheet.Dimension.End.Column) {
        $worksheet.Column($column).AutoFit()
    }

    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

    Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
} else {
    Write-Host "Save operation canceled." -ForegroundColor Yellow
}

[System.GC]::Collect()

#Disconnect
Disconnect-MicrosoftTeams | Out-Null