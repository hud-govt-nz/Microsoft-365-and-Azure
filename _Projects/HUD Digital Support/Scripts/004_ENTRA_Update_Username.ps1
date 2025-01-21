Clear-Host
Write-Host '## Change User UPN and Email Address ##' -ForegroundColor Yellow

# Requirements
#Requires -Modules Microsoft.Graph.Authentication
#Requires -Modules ExchangeOnlineManagement

# Connect to Graph and Exchange
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome

	Connect-ExchangeOnline `
	    -AppId $env:DigitalSupportAppID `
	    -Organization "mhud.onmicrosoft.com" `
	    -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
	    -ShowBanner:$false
    Write-Host "Connected" -ForegroundColor Green
        
    } catch {
	    Write-Host "Error connecting. Please check your credentials and network connection." -ForegroundColor Red
	    exit 1
        }

do {
    Write-Host ''
    $UserEmailAddress = Read-Host "Please enter the users current UserPrincipalName or email"
    $User = (Get-MgUser -UserId $UserEmailAddress).DisplayName

    Write-Host "Updating User $User.DisplayName" -ForegroundColor Cyan

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

		# Update EntraID details
		Update-MgUser -UserId $UserEmailAddress -GivenName $givenName -Surname $surname -UserPrincipalName $NewUserPrincipalName -DisplayName $displayName -MailNickname $mailNickName

		# Update Exchange Online Mailbox Primary SMTP
		Start-Sleep 5
		Set-Mailbox -Identity $NewUserPrincipalName -EmailAddresses "SMTP:$($NewUserPrincipalName)","smtp:$($UserEmailAddress)" -Name $displayName -Alias $mailAlias -WindowsEmailAddress $NewUserPrincipalName

		Write-Host "Primary mail address $($UserEmailAddress) has been updated to $($NewUserPrincipalName)" -ForegroundColor cyan
	}
	'1' {
        Disconnect-MgGraph | Out-Null
        disconnect-ExchangeOnline -Confirm:$false | Out-Null
		break
	}
}


} while ($true)