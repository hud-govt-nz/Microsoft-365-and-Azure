Clear-Host
Write-Host '## Check users Aho Employee Attributes ##' -ForegroundColor Yellow

function Convert-EmpCategory {
	param(
		[string]$Category
	)

	switch ($Category) {
		"HUD_SUBSTANTIVE_POSITION" { "Permanent" }
		"ORA_HRX_CONTRACTOR" { "Contractor" }
		"ORA_HRX_CONSULTANT" { "Consultant" }
		"HUD_EXTERNAL_SECONDMENT" { "External Secondment" }
		"HUD_FIX_TERM" { "Fixed Term" }
		"HUD_INTERNAL_SECONDMENT" { "Internal Secondment" }
		"HUD_LEAVE_WO_PAY" { "Leave Without Pay" }
		"HUD_PARENTAL_LEAVE" { "Parental Leave" }
		default { $Category }
	}
}

# Connect to MgGraph and define scope for reviewing user account
try {
	Connect-MgGraph -Scopes "Directory.Read.All","Directory.ReadWrite.All","User.Read.All","User.ReadWrite.All" | Out-Null
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

	# Return Aho Applied Start Date and Employee Category
	Write-Host ''

	try {
		$User = Get-MgBetaUser -UserId $UserPrincipalName -ErrorAction Stop
		$EmpCategory = Convert-EmpCategory -Category ($User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory | Select-Object -First 1)
		$StartDate = ($User.AdditionalProperties.extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate | Select-Object -First 1)

		# Prompt for and process an update to $EmpCategory:
		$UpdateEmpCategory = Read-Host "Current employee category is ${EmpCategory}. Enter a new employee category to update (or press Enter to keep current value):"

		if ($UpdateEmpCategory -ne '') {
			$EmpCategory = Convert-EmpCategory -Category $UpdateEmpCategory

			$Attributes = @{
				'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EmpCategory
			}

			Update-MgBetaUser -UserId $UserPrincipalName -AdditionalProperties $Attributes
		}

	} catch [Microsoft.Graph.ServiceException]{
		if ($_.Message -like "*Code: Request_ResourceNotFound*") {
			Write-Host "User '$UserPrincipalName' not found. Please check if the User Principal Name is correct and try again." -ForegroundColor Red
		} else {
			Write-Host "Error retrieving user information. Please check if the User Principal Name is correct and try again." -ForegroundColor Red
		}
		continue
	} catch {
		Write-Host "Error retrieving user information. Please check if the User Principal Name is correct and try again." -ForegroundColor Red
		continue
	}

	Write-Host "${UserPrincipalName} Employee Category is ${EmpCategory} with a start date of ${StartDate}" -ForegroundColor Green
	Write-Host "Entra User ID: $($User.id)" -ForegroundColor Green

} while ($true)
