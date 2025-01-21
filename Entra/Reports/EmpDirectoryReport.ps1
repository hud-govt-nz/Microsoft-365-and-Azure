Clear-Host
Write-Host '## Employee Directory Report ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication

# Connect to Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token        = $CollectToken.RequestMessage.Headers.Authorization.Parameter

    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

# Selected Values
$select = @(
    'id'
    'displayName'
    'mail'
    'jobTitle'
    'department'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup'
    'companyName'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'
    'manager'
    'OfficeLocation'
    'businessPhones'
    'mobilePhone'
    'assignedLicenses'
    'extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox'
    'extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox'
    'extension_56a473fa1d5b476484f306f7b06ee688_ServiceAccount'

) -join ','

# Graph API Call
$uri     = "https://graph.microsoft.com/v1.0/users?`$Filter=accountEnabled eq true and UserType eq 'Member'&`$select=$select&`$expand=manager"
$headers = @{
        "Authorization" = $Token
        "Content-Type"  = "application/json"
}

# Results
$output = @()

do {
    $req = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
    $uri = $req.'@odata.nextLink'

    foreach ($user in $req.value) {
        $output += [PSCustomObject]@{
            # Identity
            'id'                          = $user.id
            'display_name'                = $user.displayName
            'email'                       = $user.mail

            # Organisational Structure
            'job_title'                   = $user.jobTitle
            'department'                  = $user.department
            'organisational_group'        = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
            'organisation'                = $user.companyName
            'employee_type'               = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType
            'manager_id'                  = $user.manager.id
            'manager_display_name'        = $user.manager.displayName
            'manager_upn'                 = $user.manager.userPrincipalName
            'manager_job_title'           = $user.manager.jobTitle

            # Contact and Location
            'office'                      = $user.OfficeLocation
            'phone'                       = $user.businessPhones -join ','
            'mobile'                      = $user.mobilePhone
                        
            # Other
            'room_mailbox'                = $User.extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox
            'shared_mailbox'              = $User.extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox
            'Service_Account'             = $User.extension_56a473fa1d5b476484f306f7b06ee688_ServiceAccount
        }
    }
} while ($uri)
# Filtering of erroneous accounts
$output = $output | Where-Object {
    $_.Service_Account -ne "Service Account" -and
    $_.Room_Mailbox -ne "Room Mailbox" -and
    $_.Shared_Mailbox -ne "Shared Mailbox" -and
    $_.Job_Title -ne "MBIE Service Centre Analyst" -and
    ($_.employee_type -ne "Pending Worker" -or $_.employee_type -eq $null) -and
    ($_.organisation -ne "Disabled Account" -and $_.organisation -ne "Bridon Group Ltd." -and $_.organisation -ne $null)
}

# Sort the result by display name
$output = $output | Sort-Object -Property display_name

Write-Host "Open Save Dialog"

# Basic metrics
$totalUsers    = $output.Count
$enabledUsers  = $output | Where-Object { $_.'Account Enabled' -eq $true } | Measure-Object | Select-Object -ExpandProperty Count
$disabledUsers = $totalUsers - $enabledUsers

Write-Host "Total Users: $totalUsers"
Write-Host "Enabled Users: $enabledUsers"
Write-Host "Disabled Users: $disabledUsers"

$Date     = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Entra Active Staff Report"

# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog  
$SaveFileDialog.Filter   = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title    = "Save as"
$SaveFileDialog.FileName = $FileName

# Show the SaveFileDialog and get the selected file path  
$SaveFileResult = $SaveFileDialog.ShowDialog()

if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $SelectedFilePath = $SaveFileDialog.FileName
    $output | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
    
    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet    = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow    = 1
    $endRow      = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn   = $worksheet.Dimension.End.Column

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

} else {
Write-Host "Save cancelled" -ForegroundColor Yellow
}
Disconnect-MgGraph | Out-Null