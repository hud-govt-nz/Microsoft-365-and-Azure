<#
    .NAME: Get Password Expiry for hud.govt.nz user accounts
    .PURPOSE: To export a list of users who's account passwords are due to expire.
    .AUTHOR: Ashley Forde
    .VERSION: 1.0



#>

#Connect to AzureAD/MSOL



#Grab password expiry dates for user
Get-MsolUser -UserPrincipalName 'Username' |Select LastPasswordChangeTimestamp

$PasswordPolicy = Get-MsolPasswordPolicy
$UserPrincipal  = Get-MsolUser -UserPrincipalName 'Username'

$PasswordExpirationDate = $UserPrincipal.LastPasswordChangeTimestamp.AddDays($PasswordPolicy.ValidityPeriod)


#Export to .CSV




#Send Email to User list from .csv or Digital Team.


