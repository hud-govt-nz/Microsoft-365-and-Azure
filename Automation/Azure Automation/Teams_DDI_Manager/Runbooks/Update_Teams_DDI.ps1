# Define parameter block
param
(
	[Parameter(Mandatory = $true)]
	[object]$WebhookData
)

#Connect EXO
. .\TeamsLogin.ps1

#Connect to Microsoft Graph
. .\GraphLogin.ps1

# Extracting parameters from WebhookData
$WebhookBody = $WebhookData.RequestBody | ConvertFrom-Json
$User = $WebhookBody.User
$DDI = $WebhookBody.DDI

# Update and validate parameters
if ([string]::IsNullOrEmpty($User) -or [string]::IsNullOrEmpty($DDI)) {
	Write-Error "User or DDI parameter is not provided."
	exit
}

# Obtain users location
$Location = (Get-MgUser -UserId $User).OfficeLocation
$DisplayName = (Get-MgUser -UserId $User).DisplayName
$ID = (Get-MgUser -UserId $User).id

#Update EntraID
Write-Output "Updating Business Phone value in EntraID for $User"
Update-MgUser -UserId $ID -BusinessPhones $DDI

#Format DDI for application "+64XXXXXXXX"
$DDI = [string]$DDI -replace " ",''

#Check current number assignment
$currentNoAssignment = Get-CsPhoneNumberAssignment -TelephoneNumber $DDI

#Check current user assignment
$currentUserAssignment = Get-CsPhoneNumberAssignment -AssignedPstnTargetId $ID

#Check if number is already assigned to user
if ($currentNoAssignment.AssignedPstnTargetId -match $ID) {
	Write-Output "$User is already assigned to DDI: $DDI"

	# Check if the user has a phone number assigned and removes it
} elseif ($currentUserAssignment.TelephoneNumber) {
	Write-Output "Clearing existing phone number assignment for $User..."
	Remove-CsPhoneNumberAssignment -Identity $User -RemoveAll

	Start-Sleep 10

	# Assign new phone number
	Write-Output "Assigning new phone number to $User..."
	Set-CsPhoneNumberAssignment -Identity $User -PhoneNumber $DDI -PhoneNumberType DirectRouting
	Start-Sleep 10

	# Enable voicemail for the user
	Write-Output "Enabling voicemail for $User..."
	Set-CsOnlineVoicemailUserSettings -Identity $User -VoicemailEnabled $true

} else {
	# Assign new phone number
	Write-Output "Assigning new phone number to $User..."
	Set-CsPhoneNumberAssignment -Identity $User -PhoneNumber $DDI -PhoneNumberType DirectRouting
	Start-Sleep 10

	# Enable voicemail for the user
	Write-Output "Enabling voicemail for $User..."
	Set-CsOnlineVoicemailUserSettings -Identity $User -VoicemailEnabled $true
}

#Setting Dial Plan Policy based on Office Location
if ($Location -match 'Wellington') {
	Write-Output "Enabling Wellington Dial Plan Policy 'DP-04Region' for $User..."
	Grant-CsTenantDialPlan -Identity $User -PolicyName "DP-04Region"
	Grant-CsOnlineVoiceRoutingPolicy -Identity $User -PolicyName Tag:VP-Unrestricted
} else {
	Write-Output "Enabling Auckland Dial Plan Policy 'DP-09Region' for $User..."
	Grant-CsTenantDialPlan -Identity $User -PolicyName "DP-09Region"
	Grant-CsOnlineVoiceRoutingPolicy -Identity $User -PolicyName Tag:VP-Unrestricted
}

# Result
Write-Output "$DisplayName has been assigned number $DDI"

#Prevent the script from failing with a maximum of 3 allowed connections to EXO.
Get-PSSession | Remove-PSSession
