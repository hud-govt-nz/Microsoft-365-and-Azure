Clear-Host
Write-Host ''
Write-Host '## change users email address in Exchange Online ##' -ForegroundColor Yellow

# Connect to MgGraph and define scope for user account modification
Connect-MgGraph -Scopes "Directory.Read.All","Directory.ReadWrite.All","User.Read.All","User.ReadWrite.All" | Out-Null

# Connect to Exchange Online
$UserPrincipalName = Whoami /UPN
Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

# Obtain User ID
$UserEmailAddress = Read-Host "Please enter the users current UserPrincipalName or email"
$User = (Get-MgUser -UserId $UserEmailAddress).DisplayName

# Cofirm change required
$title = "Update User Email"
$sendasmsg = "Would you like to update the email and display name address for $($User)?"
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
$result = $host.ui.PromptForChoice($title,$sendasmsg,$options,0)
switch ($result) {
	'0' {
		Write-Host ''
		Write-Host "Provide new details below" -ForegroundColor Yellow
		Write-Host ''
		$givenName = Read-Host "Enter first name"
		$surname = Read-Host "Enter last name"

		$displayName = "$givenName $surname"
		$mailNickName = "$givenName$surname"
		$mailAlias = "$givenName.$surname"
		$NewUserPrincipalName = "$givenName.$surname" + "@hud.govt.nz"

		# Update Azure AD details
		Update-MgUser -UserId $UserEmailAddress -GivenName $givenName -Surname $surname -UserPrincipalName $NewUserPrincipalName -DisplayName $displayName -MailNickname $mailNickName

		# Update Exchange Online Mailbox Primary SMTP
		Start-Sleep 5
		Set-Mailbox -Identity $NewUserPrincipalName -EmailAddresses "SMTP:$($NewUserPrincipalName)","smtp:$($UserEmailAddress)" -Name $displayName -Alias $mailAlias -WindowsEmailAddress $NewUserPrincipalName

		Write-Host "Primary mail address $($UserEmailAddress) has been updated to $($NewUserPrincipalName)" -ForegroundColor cyan
	}
	'1' {
		break # Exit the loop if user selects "No"
	}
}
