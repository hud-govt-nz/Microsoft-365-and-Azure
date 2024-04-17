#Get All Groups where Welcome Email is enabled
$WelcomeEmailGroups = Get-UnifiedGroup | Where-Object { $_.WelcomeMessageEnabled -eq $True }
 
#Disable Welcome Email
ForEach($Group in $WelcomeEmailGroups) 
{
    #Disable the Group Welcome Message Email
    Set-UnifiedGroup -Identity $Group.Id -UnifiedGroupWelcomeMessageEnabled:$false
    Write-host "Welcome Email Disabled for the Group:"$Group.PrimarySmtpAddress
}