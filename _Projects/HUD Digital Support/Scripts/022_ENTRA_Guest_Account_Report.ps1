Clear-Host
Write-Host '## EntraID Guest Account Export ##' -ForegroundColor Yellow

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
    'givenName'
    'surname'
    'displayName'
    'userPrincipalName'
    'mail'
    'userType'
    'accountEnabled'
    'jobTitle'
    'department'
    'companyName'
    'City'
    'country'
    'mobilePhone'
    'ExternalUserState'
    'ExternalUserStateChangeDateTime'
    'CreationType'
    'createdDateTime'
    'signInActivity'
    'usageLocation'

) -join ','

# Graph API Call
$uri     = "https://graph.microsoft.com/v1.0/users?`$Filter=UserType eq 'Guest'&`$select=$select&`$expand=manager"
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
    
        $TimeZone = "New Zealand Standard Time"
        
        $createdDateTime                 = $user.createdDateTime
        $successfulSignIn                = $user.signInActivity.lastSuccessfulSignInDateTime
        $ExternalUserStateChangeDateTime = $user.ExternalUserStateChangeDateTime


        $NZSTcreatedDateTime                 = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($createdDateTime, [System.TimeZoneInfo]::Local.Id, $TimeZone)
        $NZSTsuccessfulSignIn                = if ($successfulSignIn) { [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($successfulSignIn, [System.TimeZoneInfo]::Local.Id, $TimeZone) } else { $null }
        $NZSTExternalUserStateChangeDateTime = if ($NZSTExternalUserStateChangeDateTime) { [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($ExternalUserStateChangeDateTime, [System.TimeZoneInfo]::Local.Id, $TimeZone) } else { $null }
        $TimeSpan                            = if ($successfulSignIn) { New-TimeSpan -Start $NZSTsuccessfulSignIn -End (Get-Date) | Select-Object -ExpandProperty Days } else { $null }
        
        $output += [PSCustomObject]@{
            # Identity
            'id'                  = $user.id
            'first_name'          = $user.givenName
            'last_name'           = $user.surname
            'display_name'        = $user.displayName
            'user_principal_name' = $user.userPrincipalName
            'email'               = $user.mail
            'user_type'           = $user.userType
            'account_enabled'     = $user.accountEnabled

            # Organisational Structure
            'job_title'    = $user.jobTitle
            'department'   = $user.department
            'organisation' = $user.companyName

            # Contact and Location
            'city'    = $user.City
            'country' = $user.country
            'mobile'  = $user.mobilePhone

            # Account
            'usage_location'                        = $user.usageLocation
            'created_date_time'                     = $NZSTcreatedDateTime
            'ExternalUserState'                     = $user.ExternalUserState
            'ExternalUserStateChangeDateTime'       = if ($NZSTExternalUserStateChangeDateTime) {$NZSTExternalUserStateChangeDateTime} else { $null }
            'CreationType'                          = $user.CreationType
            'last_successful_sign_in_date_time_nzt' = $NZSTsuccessfulSignIn
            'days_since_last_sign_in'               = $TimeSpan

        }
    }
} while ($uri)

Write-Host "Open Save Dialog"

# Basic metrics
$totalUsers    = $output.Count
$enabledUsers  = $output | Where-Object { $_.'Account Enabled' -eq $true } | Measure-Object | Select-Object -ExpandProperty Count
$disabledUsers = $totalUsers - $enabledUsers

Write-Host "Total Users: $totalUsers"
Write-Host "Enabled Users: $enabledUsers"
Write-Host "Disabled Users: $disabledUsers"

$Date     = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Entra Guest Account Export"

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