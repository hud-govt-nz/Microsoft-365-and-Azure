#Connect to Microsoft Graph
Connect-MgGraph -Identity 

# Obtain graph access token.
$CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
$Token        = $CollectToken.RequestMessage.Headers.Authorization.Parameter

# Disable PnP PowerShell update check
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

# Connect to SharePoint Online
$siteUrl = "https://mhud-admin.sharepoint.com"
Connect-PnPOnline -Url $siteUrl -ManagedIdentity

# Selected Values
$select = @(
    'id'
    'displayName'
    'mail'
    'userprincipalname'
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
            'userprincipalname'           = $user.userPrincipalName

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

# Basic metrics
$totalUsers    = $output.Count
Write-output "Total Users: $totalUsers"


# Sort the result by display name
$output = $output | Sort-Object -Property display_name | Format-Table -Property id, display_name, email, userprincipalname, job_title, department, organisational_group -AutoSize -Wrap




# Update the Organisational Group property in SharePoint Online
$output | ForEach-Object {
    $user = $_
    $user | Select-Object -Property id, display_name, email, userprincipalname, job_title, department, organisational_group
    Set-PnPUserProfileProperty -Account $user.userprincipalname -PropertyName "OrganisationalGroup" -Value $user.organisational_group

    Write-Output "Updated Organisational Group for $($user.display_name) to $($user.organisational_group)"
}


#set-PnPUserProfileProperty -Account Ashley.Forde@hud.govt.nz -PropertyName "OrganisationalGroup" -Value "Organisational Performance"

#>

#return $output