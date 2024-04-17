## Distribution Group bulk management

# Connection
Connect-ExchangeOnline -ShowBanner:$false

# Return DL
$OldDistName = "DL - Te Kāhui Māori Housing Supply and Delivery"
#Get-DistributionGroup -Identity $OldDistName | Select-Object Name, Alias, DisplayName, PrimarySmtpAddress,EmailAddresses, managedby | Format-list 

# Rename distribution group and SMTP Addresses
$NewDistName = "DL - SDP - Māori Housing Supply and Delivery Team"
$NewDistAlias = "DL-SDP-MāoriHousingSupplyandDeliveryTeam"
$NewDistEmail = "DL-SDP-MāoriHousingSupplyandDeliveryTeam@hud.govt.nz"
 
# Change distribution group Alias and DisplayName
Set-DistributionGroup -Identity $OldDistName -Name $NewDistName -Alias $NewDistAlias -DisplayName $NewDistName
 
# Change distribution group EmailAddresses
Set-DistributionGroup -Identity $NewDistName -PrimarySmtpAddress $NewDistEmail

# Get Renamed Distribution Group
Get-DistributionGroup -Identity $NewDistName | Select-Object Name, Alias, DisplayName, PrimarySmtpAddress,EmailAddresses, managedby | Format-list 

# Get Distribution Group Members

$DLGPMEMBR = Get-DistributionGroupMember -ResultSize Unlimited "DL - Strategy, Insight and Governance (SIG)" | Sort -Property department | Select-Object DisplayName, Department, primarySMTPAddress 
$DLGPMEMBR | Export-Excel C:\HUD\06_Reporting\DL_SIGMEMBERS.xlsx -AutoSize -AutoFilter -WorksheetName 'Members' -FreezeTopRow -BoldTopRow

# Add members to Distribution Group

Foreach ($i in $Email) {
    #Add-DistributionGroupMember -Identity $NewDistName -Member $I.UserPrincipalName -Confirm:$false
    Remove-DistributionGroupMember -Identity "DL - Strategy, Insight and Governance (SIG)" -Member $I.UserPrincipalName -Confirm:$false
}