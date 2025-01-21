Clear-Host
Write-host "## TEAMS: Owner and Member Report ##" -ForegroundColor Yellow

# Connect to Teams
try {
    Connect-MicrosoftTeams -TenantId $env:DigitalSupportTenantID -CertificateThumbprint $env:DigitalSupportCertificateThumbprint -ApplicationId $env:DigitalSupportAppID
    Write-Host "Connected to Microsoft Teams."
        
    } catch {
	    Write-Host "Error connecting to Microsoft Teams. Please check your credentials and network connection." -ForegroundColor Red
	    exit 1
}

# Get all Teams
$AllTeams = Get-Team | Select-Object DisplayName, GroupId, MailNickName, Visibility
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

Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Teams Owners and Members"

# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms
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
    $Export = $TeamData | Select-Object TeamID,TeamName,TeamType,MailAlias,TeamOwners,TeamMembers | Sort-Object TeamName
    $Export | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow

    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow = 1
    $endRow = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn = $worksheet.Dimension.End.Column

    # Set horizontal alignment to left for all cells in the used range
    for ($col = $startColumn; $col -le $endColumn; $col++) {
        for ($row = $startRow; $row -le $endRow; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
            }
        }

    # Autosize columns if needed
    foreach ($column in $worksheet.Dimension.Start.Column.$worksheet.Dimension.End.Column) {
        $worksheet.Column($column).AutoFit()
        }

    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

    Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
    }

#Disconnect
Disconnect-MicrosoftTeams | Out-Null