
#Connect to Exchange Online (note, this can also be run for exchange on-premise users)
Connect-ExchangeOnline

#User Mailboxes
$UserMailboxes = get-Mailbox | Get-user -ResultSize unlimited -Filter "(RecipientType -eq 'UserMailbox')"
$Users = $UserMailboxes.Name


#Update
ForEach ($User in $Users) {Set-User -Identity $User -RemotePowerShellEnabled $false}
#$UserMailbox | foreach {Set-User -Identity $_.WindowsEmailAddress -RemotePowerShellEnabled $false}

#Show Status
$UserMailboxes | Select-Object Name, RecipientType, RemotePowershellEnabled


<#SINGLE USER EXAMPLE

Set-User -Identity "Test User" -RemotePowerShellEnabled $False

#>