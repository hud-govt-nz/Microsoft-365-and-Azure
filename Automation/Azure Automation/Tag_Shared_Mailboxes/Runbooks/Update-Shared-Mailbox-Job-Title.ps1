#Connect EXO
. .\EXOLogin.ps1

#Connect to Microsoft Graph
. .\MgGraphLogin.ps1

#Get all shared mailboxes
$SMBX = (Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited).UserPrincipalname
$RMBX = (Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited).UserPrincipalname

#Set all SMBS on Jobtitle Shared Mailbox and set AccountEnabled on false
foreach ($SMB in $SMBX) {
	$ATB1 = @{ 'extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox' = "Shared Mailbox" }
	Update-MgUser -UserId $SMB -JobTitle "Shared Mailbox" -AdditionalProperties $ATB1
}

foreach ($RMB in $RMBX) {
	$ATB2 = @{ 'extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox' = "Room Mailbox" }
	Update-MgUser -UserId $RMB -JobTitle "Room Mailbox" -AdditionalProperties $ATB2
}


#Prevent the script from failing with a maximum of 3 allowed connections to EXO.
Get-PSSession | Remove-PSSession
