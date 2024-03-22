# Connect to Microsoft Graph with user read/write permissions
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome | Out-Null

Add-Type -AssemblyName System.Windows.Forms

# Import data from Excel
$fd = New-Object System.Windows.Forms.OpenFileDialog
$fd.Filter = "Excel Workbook|*.xlsx"
$fd.InitialDirectory = [Environment]::GetFolderPath('Desktop')
$fd.ShowDialog() | Out-Null
$Import = Import-Excel -Path $fd.FileName

# Filter for rows with non-null 'First Name'
$Users = $Import | Where-Object { $_.'First Name' -ne $null }

# Set Password Profile
$PasswordProfile = @{
    Password                             = "b48a4392-b31f-265c-87d4-a351fcda5efb"
    ForceChangePasswordNextSignIn        = $true
    }

# Loop through each user and create new user if from Wellington
foreach ($User in $Users) {
    if ($User.City -eq "Wellington") {$PostalCode = "6011" 
        } elseif ($User.City -eq "Auckland") {$PostalCode = "1010" 
            }
    try {
        # Set user details
        $DisplayName = "$($User.'First Name') $($User.'Last Name')"
        $UPN = "$($User.'First Name' -replace '\s','').$($User.'Last Name' -replace '\s','')@hud.govt.nz"
        $MailNickName = "$($User.'First Name' -replace '\s','')$($User.'Last Name' -replace '\s','')"
        $StartDate = $User.'Start Date'.ToString("dd/MM/yyyy")
    
        # Create new user
        $NewUser = New-MgUser `
            -GivenName $User.'First Name' -Surname $User.'Last Name' -DisplayName $DisplayName -UserPrincipalName $UPN -MailNickname $MailNickName `
            -JobTitle $User.'Job Title' -Department $User.Department -OfficeLocation $User.City `
            -City $User.City -StreetAddress $User.'Street Address' -PostalCode $PostalCode -Country "New Zealand" -State "NZ" `
            -UsageLocation "NZ" -PreferredLanguage "en-NZ" -PasswordProfile $PasswordProfile `
            -PasswordPolicies "DisablePasswordExpiration" -AccountEnabled 
    
        # Assign manager
        $Managerid = (Get-MgUser -UserId $User.Manager).id
        $NewManager = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$Managerid"}
        Set-MgUserManagerByRef -UserId $NewUser.Id -BodyParameter $NewManager

        # Assign additional properties
        $additionalProperties = @{
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'             = $User.'Employee Type'
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'       = $User.'Employee Category'
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'                = $StartDate
            'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup'      = $User.'Organisational Group'
            }
        Update-MgBetaUser -UserId $NewUser.Id -AdditionalProperties $additionalProperties

        # Capture user details
        $result = Get-mgbetauser -UserId $NewUser.Id 
        $Manager = (Get-MgUser -UserId $User.Manager).UserPrincipalName

        $Output = [pscustomobject]@{
            ObjectID                    = $result.Id
            GivenName                   = $result.GivenName
            Surname                     = $result.Surname
            DisplayName                 = $result.DisplayName
            UserPrincipalName           = $result.UserPrincipalName
            JobTitle                    = $result.JobTitle
            Department                  = $result.Department
            'Organisational Group'      = $result.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
            Office                      = $result.OfficeLocation
            Address                     = $Result.StreetAddress
            'Employee Type'             = $result.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType
            'Employee Category'         = $result.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory
            'Start Date'                = $result.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate
            Manager                     = $Manager
            }
    $Output
            
        } catch {
            Write-Host ("Failed to create the account for {0}. Error: {1}" -f $DisplayName, $_.Exception.Message) -ForegroundColor Red
            }
    }

