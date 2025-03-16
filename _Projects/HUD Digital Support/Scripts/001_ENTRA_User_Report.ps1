<#
.SYNOPSIS
    Exports detailed user information from Microsoft Entra ID (Azure AD) to Excel.

.DESCRIPTION
    This script connects to Microsoft Graph API and retrieves comprehensive user information 
    from Entra ID, including user details, organizational structure, contact information,
    and account status. The data is exported to an Excel file with formatted columns.

.NOTES
    Version:        1.0
    Author:         Digital Support Team
    Creation Date:  2024
    
.REQUIREMENTS
    - Microsoft.Graph.Authentication module
    - ImportExcel module
    - Environment variables:
        * DigitalSupportAppID
        * DigitalSupportTenantID
        * DigitalSupportCertificateThumbprint

.OUTPUTS
    Excel file containing user information with the following details:
    - Identity information
    - Organizational structure
    - Contact and location details
    - Account status and properties
    - License information
#>

Clear-Host
Write-Host '## EntraID User Account Export ##' -ForegroundColor Yellow

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

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users?$select"accountenabled" -ContentType "txt" -OutputType HttpResponseMessage
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
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup'
    'companyName'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'
    'manager'
    'OfficeLocation'
    'streetAddress'
    'City'
    'postalCode'
    'state'
    'country'
    'businessPhones'
    'mobilePhone'
    'createdDateTime'
    'signInActivity'
    'signInActivity'
    'usageLocation'
    'passwordPolicies'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime'
    'assignedLicenses'
    'SecurityIdentifier'
    'extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox'
    'extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox'
    'extension_56a473fa1d5b476484f306f7b06ee688_ServiceAccount'

) -join ','

# Graph API Call
$uri     = "https://graph.microsoft.com/v1.0/users?`$Filter=UserType eq 'Member'&`$select=$select&`$expand=manager"
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

        $NZSTcreatedDateTime                 = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($createdDateTime, [System.TimeZoneInfo]::Local.Id, $TimeZone)
        $NZSTsuccessfulSignIn                = if ($successfulSignIn) { [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($successfulSignIn, [System.TimeZoneInfo]::Local.Id, $TimeZone) } else { $null }
        $TimeSpan                            = if ($successfulSignIn) { New-TimeSpan -Start $NZSTsuccessfulSignIn -End (Get-Date) | Select-Object -ExpandProperty Days } else { $null }

        $output += [PSCustomObject]@{
            # Identity
            'id'                          = $user.id
            'first_name'                  = $user.givenName
            'last_name'                   = $user.surname
            'display_name'                = $user.displayName
            'user_principal_name'         = $user.userPrincipalName
            'email'                       = $user.mail
            'user_type'                   = $user.userType
            'account_enabled'             = $user.accountEnabled

            # Organisational Structure
            'job_title'                   = $user.jobTitle
            'department'                  = $user.department
            'organisational_group'        = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
            'organisation'                = $user.companyName
            'employee_type'               = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType
            'employee_category'           = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory
            'manager_display_name'        = $user.manager.displayName
            'manager_upn'                 = $user.manager.userPrincipalName
            'manager_job_title'           = $user.manager.jobTitle

            # Contact and Location
            'office'                      = $user.OfficeLocation
            'address'                     = $user.streetAddress
            'city'                        = $user.City
            'postal_code'                 = $user.postalCode
            'state'                       = $user.state
            'country'                     = $user.country
            'phone'                       = $user.businessPhones -join ','
            'mobile'                      = $user.mobilePhone

            # Account
            'created_date_time'                      = $NZSTcreatedDateTime
            'last_successful_sign_in_date_time_nzt' = $NZSTsuccessfulSignIn
            'days_since_last_sign_in'               = $TimeSpan
            'usage_location'                         = $user.usageLocation
            'password_policies'                      = $user.passwordPolicies
            'start_date'                             = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate
            'leave_date'                             = $user.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime
            'e5_license'                             = if ($user.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06") { $true } else { $false }
            'no_licenses'                            = if ($user.assignedLicenses.count -eq 0) { $true } else { $false }
                        
            # Other
            'security_identifier'         = $user.SecurityIdentifier
            'room_mailbox'                = $User.extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox
            'shared_mailbox'              = $User.extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox
            'Service_Account'             = $User.extension_56a473fa1d5b476484f306f7b06ee688_ServiceAccount
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
$FileName = "Entra All Users Export"

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