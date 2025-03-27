<#
.SYNOPSIS
    Activate PIM roles via Graph PowerShell.

.DESCRIPTION
    Script can be used to activate PIM eligible assignments using native Graph PowerShell cmdlets.
    The script dynamically retrieves and displays only the roles that the current user is eligible for.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Activate-PIM-Assignment.ps1

.NOTES
    - AUTHOR : Ashley Forde
    - Version: 3.0
    - Date   : 27 Mar 2025
    - Notes  : Updated to dynamically show only eligible roles for the current user instead of using a static list
#>

#Requires -Modules Microsoft.Graph.Identity.Governance

Clear-Host

# Connect to Graph 
Connect-MgGraph -NoWelcome

# Obtain current user EntraID ObjectID
$context     = Get-MgContext
$currentUser = (Get-MgUser -UserId $context.Account).id

# Display header
Write-Host '## Elevate PIM roles ##' -ForegroundColor Yellow

# Get all roles that the current user is eligible for
Write-Host "Retrieving your eligible roles..." -ForegroundColor Cyan
$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentuser'"

if (-not $myRoles) {
    Write-Host "You don't have any eligible roles to activate." -ForegroundColor Red
    exit
}

# Extract role display names from the eligible roles
$ToActivate = $myRoles | ForEach-Object { $_.RoleDefinition.DisplayName } | Sort-Object

Write-Host "Found $($ToActivate.Count) eligible role(s) for your account." -ForegroundColor Green

# Let user select from their eligible roles
$SelectedRoles = $ToActivate | Out-GridView -Title "Select Role(s) to Activate" -PassThru

if (-not $SelectedRoles) {
    Write-Host "No roles selected. Exiting..." -ForegroundColor Yellow
    exit
}

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

# Process each selected role
foreach ($value in $SelectedRoles) {
	# Get role displayName and ID from the already retrieved eligible roles
	$Role = $myRoles | Where-Object { $_.RoleDefinition.DisplayName -eq $value } | Select-Object -ExpandProperty RoleDefinition | Select-Object DisplayName, Id

	# Check for existing active assignments
	$existingAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "principalId eq '$currentuser' and roleDefinitionId eq '$($Role.id)'"
    
	if ($existingAssignments) {
		Write-Host "Note: Found existing assignment for role '$value'" -ForegroundColor Yellow
		$latestAssignment = $existingAssignments | Sort-Object -Property CreatedDateTime -Descending | Select-Object -First 1
        
		if ($latestAssignment.AssignmentType -eq 'Activated') {
			$expirationTime = $latestAssignment.ScheduleInfo.Expiration.EndDateTime
			Write-Host "Role '$value' is already active until $expirationTime" -ForegroundColor Cyan
			$extendChoice = Read-Host "Would you like to extend this role assignment? (Yes/No)"
			if ($extendChoice -ne 'Yes') {
				continue
			}
			Write-Host "Proceeding to extend role assignment..." -ForegroundColor Green
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

	# Get the specific role details from the list of eligible roles we already retrieved
	$myRole = $myRoles | Where-Object { $_.RoleDefinitionId -eq "$($Role.id)" }

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
