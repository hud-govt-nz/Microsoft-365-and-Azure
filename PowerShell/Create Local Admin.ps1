# New Account
New-LocalUser -Name "administratoraccounthere" -Password (ConvertTo-SecureString "superstrongpasswordhere" -AsPlainText -Force) -Description "Administrator"
Add-LocalGroupMember -Group "Administrators" -Member "HUDAdmin"


# Update password
$account = Get-LocalUser -Name "NewAdminAccount"
$account | Set-LocalUser -Password (ConvertTo-SecureString "NewPassword" -AsPlainText -Force)
