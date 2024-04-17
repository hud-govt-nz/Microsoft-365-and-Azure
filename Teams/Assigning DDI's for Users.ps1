<#

This script is used after a DDI number has been assiged on the MASTER DDI Spreadsheet
#>


#Connect-AzureAD
#Connect-MicrosoftTeams

#Get Users UPN
$User = Read-Host "Please provide the users User Principal Name"
$Location = Read-Host "Please provide location, either Auckland or Wellington"
$DDI = Read-Host "Please Provide the users DDI phone number. Format must be +64XXXXXXXX"

#Assign Number in Azure
Set-AzureADUser -ObjectId $User -TelephoneNumber $DDI 
$UPN = [string](Get-AzureADUser -ObjectId $User).UserPrincipalName

#Assign number in Microsoft Teams

Set-CsPhoneNumberAssignment -Identity $UPN -PhoneNumber $DDI -PhoneNumberType DirectRouting      
Set-CsOnlineVoicemailUserSettings -Identity $UPN -VoicemailEnabled $true       

if ($Location -eq 'Wellington') {
    Grant-CsTenantDialPlan -Identity $UPN  -PolicyName "DP-04Region" 
    Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN  -PolicyName Tag:VP-Unrestricted 

} else {
    Grant-CsTenantDialPlan -Identity $UPN -PolicyName "DP-09Region"
    Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName Tag:VP-Unrestricted

}