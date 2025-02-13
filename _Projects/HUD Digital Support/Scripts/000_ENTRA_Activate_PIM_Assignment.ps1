<#
.SYNOPSIS
    Activate PIM roles via Graph PowerShell.

.DESCRIPTION
    Script can be used to activate PIM eligable assignments using native Graph Powershell cmdlets.  

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Enable_PIM_Assignment.ps1
    powershell.exe -executionpolicy bypass -file .\Enable_PIM_Assignment.ps1

.NOTES
    - AUTHOR : Ashley Forde
    - Version: 2.0
    - Date   : 12 Oct 2023
    - Notes  : 
#>

#Requires -Modules Microsoft.Graph.Identity.Governance

Clear-Host

# Connect to Graph 
Connect-MgGraph -NoWelcome

# Obtain current user EntraID ObjectID
$context     = Get-MgContext
$currentUser = (Get-MgUser -UserId $context.Account).id

# Select which role group you wish to activate
#Clear-Host
Write-Host '## Elevate PIM roles ##' -ForegroundColor Yellow

$ToActivate = @(
	'Application Administrator',
	'Attack Payload Author',
	'Attack Simulation Administrator',
	'Authentication Administrator',
	'Authentication Policy Administrator',
	'Azure Information Protection Administrator',
	'Cloud App Security Administrator',
	'Cloud Device Administrator',
	'Compliance Administrator',
	'Compliance Data Administrator',
	'Conditional Access Administrator',
	'Exchange Administrator',
	'Exchange Recipient Administrator',
	'Global Administrator',
	'Guest Inviter',
	'Helpdesk Administrator',
	'Intune Administrator',
	'License Administrator',
	'Office Apps Administrator',
	'Privileged Authentication Administrator',
	'Privileged Role Administrator',
	'Security Administrator',
	'Security Operator',
	'Security Reader',
	'Sharepoint Administrator',
	'Teams Administrator',
	'Teams Communications Administrator',
	'User Administrator'
)

$SelectedRoles = $ToActivate | Out-GridView -Title "Select Role(s) to Activate" -PassThru

# Write the selected roles to the host as a comma-separated list
Write-Host "Selected Roles:" -ForegroundColor Green
Write-Host ($SelectedRoles -join ', ') -ForegroundColor Cyan

# Provide justification to activate roles
$Justification = Read-Host "Provide a justification"

# Ask the user if they want to set a custom duration
$UserChoice = Read-Host "Would you like to set a custom duration? (Yes/No) [Default is No]"

# Default to 'No' if the user presses Enter without typing anything
if ([string]::IsNullOrWhiteSpace($UserChoice)) {
	$UserChoice = 'No'
}

#$Value = "User Administrator"
foreach ($value in $SelectedRoles) {
	# Get role displayName and ID
	$Role = Get-MgDirectoryRoleTemplate | Where-Object { $value -contains $_.DisplayName } | Select-Object DisplayName, Id

	# Check for existing active assignments
	$existingAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "principalId eq '$currentuser' and roleDefinitionId eq '$($Role.id)'"
    
	if ($existingAssignments) {
		Write-Host "Note: Found existing assignment for role '$value'" -ForegroundColor Yellow
		$latestAssignment = $existingAssignments | Sort-Object -Property CreatedDateTime -Descending | Select-Object -First 1
        
		if ($latestAssignment.AssignmentType -eq 'Activated') {
			$expirationTime = $latestAssignment.ScheduleInfo.Expiration.EndDateTime
			Write-Host "Role '$value' is already active until $expirationTime" -ForegroundColor Cyan
			$extendChoice = Read-Host "Would you like to extend this role assignment? (Yes/No)"
			if ($extendChoice -eq 'Yes') {
				# Continue with the activation process which will create a new assignment
				Write-Host "Proceeding to extend role assignment..." -ForegroundColor Green
			} else {
				continue
			}
		}
	}

	# Policy Assignment to role
	$PolicyAssignment = Get-MgPolicyRoleManagementPolicyAssignment -Filter "scopeId eq '/' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '$($Role.Id)'"

	# Retrieve Rule (specific to Expiration_EndUser_Assignment but can be other rules if required)
	$Rule = Get-MgPolicyRoleManagementPolicyRule -UnifiedRoleManagementPolicyId $PolicyAssignment.PolicyId | Where-Object { $_.id -eq 'Expiration_EndUser_Assignment' } | Select-Object Id -ExpandProperty AdditionalProperties

	# Max duration role is allowed to be activated for when elevating via PIM
	$MaxRoleDuration = $Rule["maximumDuration"]

	# Parse the maximum duration into an integer for comparison
	$MaxDurationHours = [int]$MaxRoleDuration.SubString(2,$MaxRoleDuration.Length - 3)

	# If the user enters 'Yes', prompt for a custom duration; otherwise, use the max duration
	if ($UserChoice -eq 'Yes') {
		# Initialize a variable to hold the user-selected duration
		$SelectedDurationHours = 0

		# Prompt the user to enter a duration, repeating until a valid duration is entered    
		do {
			$UserInput    = Read-Host "Enter a duration for the assignment $($value) in hours (1-$MaxDurationHours)"
			$output       = 0
			$parseSuccess = [int]::TryParse($UserInput,[ref]$output)
			if ($parseSuccess -and $output -gt 0 -and $output -le $MaxDurationHours) {
				$SelectedDurationHours = $output
				$FutureDate            = (Get-Date).AddHours($output)
			} else {
				Write-Host "Invalid input. Please enter a number between 1 and $MaxDurationHours."
			}
		} while ($SelectedDurationHours -le 0 -or $SelectedDurationHours -gt $MaxDurationHours)
	} else {
		# User chose 'No' or provided an invalid response, so use the max duration
		$SelectedDurationHours = $MaxDurationHours
		$FutureDate            = (Get-Date).AddHours($MaxDurationHours)
	}

	# Convert the selected duration back to the format expected by the script 
	$SelectedDuration = "PT${SelectedDurationHours}H"

	# Format date/time for result array.
	$FormattedFutureDate = $FutureDate.ToString("dd-MMM-yy hh:mm tt")

	# Get all available roles
	$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentuser'"

	# Get SharePoint admin role info
	$myRole = $myroles | Where-Object { $_.RoleDefinitionid -eq "$($Role.id)" }

	# Setup parameters for activation
	$params = @{
		Action           = "selfActivate"
		PrincipalId      = $myRole.PrincipalId
		RoleDefinitionid = $myRole.RoleDefinitionid
		DirectoryScopeId = $myRole.DirectoryScopeId
		Justification    = "$Justification"
		ScheduleInfo     = @{
			StartDateTime = Get-Date
			Expiration    = @{
				Type     = "AfterDuration"
				Duration = "$SelectedDuration"
			}
		}
		TicketInfo = @{
			TicketNumber = $null
			TicketSystem = "$null"
		}
	}
	try {
		# Activate the role
		$Activation = New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params -ErrorAction Stop

		$Result = @()

		$Result += [pscustomobject]@{
			Role                                   = $value
			UnifiedRoleAssignmentScheduleRequestId = $Activation.id
			Justification                          = $Activation.Justification
			Status                                 = $Activation.Status
			EntraRoleID                            = $Activation.RoleDefinitionid
			Expires                                = $FormattedFutureDate
		}
		$Result

	} catch {
		if ($_.Exception.Message -like "*role assignment already exists*") {
			Write-Host "Role '$value' is already assigned or activation is in progress." -ForegroundColor Yellow
		} else {
			Write-Host "Error activating $($value) role: $($_.Exception.Message)" -ForegroundColor Red
		}
		continue
	}

}