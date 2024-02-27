Clear-Host
Write-host "## Microsoft Teams: Owner and Member Report ##" -ForegroundColor Yellow

# Connect to Microsoft Teams
Connect-MgGraph -NoWelcome | Out-Null
$UPN = (Get-MgContext).Account
Connect-MicrosoftTeams -AccountId $UPN -Confirm:$false | Out-Null

# Get all Teams
$AllTeams = Get-Team
$TeamData = @()

# Initialize progress bar variables
$TotalTeams = $AllTeams.Count
$CurrentTeam = 0

Foreach ($Team in $AllTeams)
{
    # Update progress bar
    $CurrentTeam++
    $ProgressStatus = "Processing $($Team.DisplayName) ($CurrentTeam of $TotalTeams)"
    Write-Progress -Activity "Collecting Teams Data" -Status $ProgressStatus -PercentComplete (($CurrentTeam / $TotalTeams) * 100)

    # Collect Team Data
    $TeamObjectID = $Team.GroupId.ToString()
    $TeamData += [PSCustomObject] @{
        TeamName = $Team.DisplayName
        TeamID =  $TeamObjectID
        MailAlias = $Team.MailNickName
        TeamType =  $Team.Visibility
        TeamOwners = (Get-TeamUser -GroupId $TeamObjectID | Where-Object {$_.Role -eq 'Owner'}).Name -join '; '
        TeamMembers = (Get-TeamUser -GroupId $TeamObjectID | Where-Object {$_.Role -eq 'Member'}).Name -join '; '
    }
}

# Complete the progress bar
Write-Progress -Activity "Collecting Teams Data" -Completed

# Get Date for File Naming
$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Teams Owners and Members"

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
    $Export = $TeamData | Select-Object TeamID,TeamName,TeamType,MailAlias,TeamOwners,TeamMembers | Sort-Object TeamName
    $Export | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow

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