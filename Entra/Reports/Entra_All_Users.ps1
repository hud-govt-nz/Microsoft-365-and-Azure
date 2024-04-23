# Connect to Microsoft Graph PowerShell
Write-Host "Connecting to Microsoft Graph"
$Scopes = @('AuditLog.Read.All','Directory.Read.All','Organization.Read.All','User.Read','User.Read.All',"UserAuthenticationMethod.Read.All")
Connect-MgGraph -Scopes $Scopes -NoWelcome | out-null

# Obtain Last Sign Date Time (Non Standard property value)
Write-Host "Collecting user information..." -ForegroundColor yellow

# Define the Get-AllUsers function
Function Get-AllUsers {
    param (
        [Parameter(Mandatory)]
        [bool]
        $IncludeGuests
    )
    
    process {
        # Retrieve users using the Microsoft Graph API with property
        $propertyParams = @{
            All            = $true
            #Select              = '*'
            Property            = 'SignInActivity'             
            ExpandProperty      = 'manager'
        }
            if ($IncludeGuests) {
                # Remove the filter
                $propertyParams.Remove('Filter')
            } else {
                # Keep the filter
                $propertyParams['Filter'] = "userType eq 'member'"
            }

        $users = Get-MgBetaUser @propertyParams
        $totalUsers = $users.Count

        # Initialize progress counter
        $progress = 0

        # Collect and loop through all users
        foreach ($index in 0..($totalUsers - 1)) {
            $user = $users[$index]

            # Update progress counter
            $progress++
            
            # Calculate percentage complete
            $percentComplete = ($progress / $totalUsers) * 100

            # Define progress bar parameters
            $progressParams = @{
                Activity        = "Processing Users"
                Status          = "User $($index + 1) of $totalUsers - $($user.userPrincipalName) - $($percentComplete -as [int])% Complete"
                PercentComplete = $percentComplete
            }
            
            # Display progress bar
            Write-Progress @progressParams

            if ($null -ne $User.SignInActivity -and $null -ne $User.SignInActivity.LastSignInDateTime) {
                # Convert SignInActivity to NZT
                $NZT_SignInActivity = $User.SignInActivity.LastSignInDateTime.ToUniversalTime().ToLocalTime()
                
                # Update the user object with NZT SignInActivity
                $User.SignInActivity.LastSignInDateTime = $NZT_SignInActivity
            }
        
            # Create an object to store user properties
            $userObject = [PSCustomObject]@{
                "ID"                          = $user.id
                "First name"                  = $user.givenName
                "Last name"                   = $user.surname
                "Display name"                = $user.displayName
                "User principal name"         = $user.userPrincipalName
                "Email address"               = $user.mail
                "Job title"                   = $user.jobTitle
                "Manager display name"        = $User.Manager.AdditionalProperties.displayName
                "Manager user principal name" = $User.Manager.AdditionalProperties.userPrincipalName
                "Department"                  = $user.department
                "Company"                     = $user.companyName
                "Office"                      = $user.officeLocation
                "Mobile"                      = $user.mobilePhone
                "Phone"                       = $user.businessPhones -join ','
                "Street"                      = $user.streetAddress
                "City"                        = $user.city
                "Postal code"                 = $user.postalCode
                "State"                       = $user.state
                "Country"                     = $user.country
                "User type"                   = $user.userType
                "Account status"              = if ($user.accountEnabled) { "enabled" } else { "disabled" }
                "Account Created on"          = $user.createdDateTime
                "Last log in"                 = if ($user.SignInActivity.LastSignInDateTime) { $NZT_SignInActivity } else { "No log in" }
			    "M365E3"                      = if ($user.assignedLicenses.skuid -eq "05e9a617-0261-4cee-bb44-138d3ef5d965") { $true } else { $false }
				"M365E5"                      = if ($user.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06") { $true } else { $false }
				"NoLicense"                   = if ($user.assignedLicenses.count -eq 0) { $true } else { $false }
				"UsageLocation"               = $user.usageLocation
				"SID"						  = $user.SecurityIdentifier
				"passwordPolicies"            = $User.passwordPolicies
			    "StartDate" 				  = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate
				"LeaveDate"                   = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime
				"EmployeeType"                = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType
				"Employee Category"           = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory
				"OrgGroup"                    = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup
				"RoomMailbox"                 = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox
				"SharedMailbox"               = $User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox
                "EmployeeOrgData - Division"  = $user.EmployeeOrgData.Division
}
            
            # Output the user object
            $userObject
        }
    }
}


$Query = Read-Host "Do you want to include 'Guest' accounts? (y/n)"

if ($Query -eq "y") {
	$results = Get-AllUsers -IncludeGuests $true
} else {
	$results = Get-AllUsers -IncludeGuests $false
}


Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Entra All Users Export"

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
	$results | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
    
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
    foreach ($column in $worksheet.Dimension.Start.Column..$worksheet.Dimension.End.Column) {
        $worksheet.Column($column).AutoFit()
    }
    
    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

	Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
} else {
	Write-Host "Save cancelled" -ForegroundColor Yellow
}
Clear