Clear-Host
Write-host "## TEAMS: User Access Report ##" -ForegroundColor Yellow

# Connect to Teams
try {
    Connect-MicrosoftTeams -TenantId $env:DigitalSupportTenantID -CertificateThumbprint $env:DigitalSupportCertificateThumbprint -ApplicationId $env:DigitalSupportAppID
    Write-Host "Connected to Microsoft Teams."
        
} catch {
    Write-Host "Error connecting to Microsoft Teams. Please check your credentials and network connection." -ForegroundColor Red
    exit 1
}

$UserEmail = Read-host "Please enter UPN"

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
            Write-Progress -Activity "Processing Channels" -Status "Processing channel $CurrentChannelCount of $TotalChannels $channelDispName" -PercentComplete (($CurrentChannelCount / $TotalChannels) * 100)

            try {
                $ChannelUsers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $channelDispName
            } catch {
                Write-Host "Error retrieving users for channel $channelDispName in team $TeamName. $_" -ForegroundColor Red
                continue
            }

            # Process the ChannelUsers data
            foreach ($User in $ChannelUsers) {
                if ($User.User -match $UserEmail) {
                    $AllData += [PSCustomObject]@{
                        TeamName = $TeamName
                        Name = $User.Name
                        UPN = $User.User
                        Role = $null
                        Channel = $channelDispName
                        ChannelRole = $User.Role
                    }
                }
            }
        }
    }
}

# Clear the progress bar
Write-Progress -Activity "Processing Completed" -Completed

Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "User Access in Teams"

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
    $Export = $AllData | Select-Object TeamName, Name, UPN, Role, Channel, ChannelRole | Sort-Object TeamName
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