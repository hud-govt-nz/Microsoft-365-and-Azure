Clear-Host
Write-Host '## Update Aho Employee Attributes ##' -ForegroundColor Yellow

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
        
    } catch {
		Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
		exit 1
}

do {
	# Employee Status
	Write-Host ''
	$UserPrincipalName = Read-Host "Enter the User Principal Name of the user (or 'q' to quit)"

	if ($UserPrincipalName -eq 'q') {
		break
	}

	# Return User details including Aho additional properties
	Write-Host ''

	try {
        # Capture user details
        $result = Get-mgbetauser -UserId $UserPrincipalName -ErrorAction Stop
        $Manager = (Get-MgUser -UserId $UserPrincipalName -ExpandProperty Manager | Select-Object @{Name = 'Manager'; Expression = { $_.Manager.AdditionalProperties.mail } }).manager

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

		$UpdateOption = Read-Host "Do you wish to change the user's Employee Type & Employee Category, or Organisational Group? (ETC, OG, or press Enter to keep current values):"

		if ($UpdateOption -ne '') {
			switch ($UpdateOption.ToUpper()) {
				'ETC' {
					$updateEmpType = Read-Host 'Enter a new employee type to update (or press Enter to keep current value):'
					$updateEmpCategory = Read-Host 'Enter a new employee category to update (or press Enter to keep current value):'

					if ($updateEmpType -ne '') {
						$attributes = @{
							'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType' = $updateEmpType
							'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $updateEmpCategory
						}

						Update-MgBetaUser -UserId $UserPrincipalName -AdditionalProperties $attributes
						Write-Host "User: $UserPrincipalName`nEmployee Type: $updateEmpType`nEmployee Category: $updateEmpCategory" -ForegroundColor Green
					}
				}
				'OG' {
					$updateOrgGroup = Read-Host 'Enter a new organisational group to update (or press Enter to keep current value):'
					if ($updateOrgGroup -ne '') {
						$attributes = @{
							'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = $updateOrgGroup
						}

						Update-MgBetaUser -UserId $UserPrincipalName -AdditionalProperties $attributes
						Write-Host "User: $UserPrincipalName`nOrganisational Group: $updateOrgGroup" -ForegroundColor Green
					}
				}
			}
		}
        } catch [Microsoft.Graph.ServiceException]{
            if ($_.Message -like "*Code: Request_ResourceNotFound*") {
                Write-Host "User '$UserPrincipalName' not found. Please check if the User Principal Name is correct and try again." -ForegroundColor Red
                } else {
                    Write-Host "Error retrieving user information. Please check if the User Principal Name is correct and try again." -ForegroundColor Red
                    }   
            continue
            }
} while ($true)
