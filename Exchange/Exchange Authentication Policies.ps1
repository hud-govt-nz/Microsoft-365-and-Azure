New-AuthenticationPolicy -Name "ExchangeOnlineDeskless" -AllowBasicAuthActiveSync:$false -AllowBasicAuthImap:$false -AllowBasicAuthPop:$false                                                                                                                                                                                                                                                             
New-AuthenticationPolicy -Name "ExchangeOnlineEssentials" -AllowBasicAuthActiveSync:$false -AllowBasicAuthImap:$false -AllowBasicAuthPop:$false     
New-AuthenticationPolicy -Name "ExchangeOnlineEnterprise" -AllowBasicAuthActiveSync:$false -AllowBasicAuthImap:$false -AllowBasicAuthPop:$false     
New-AuthenticationPolicy -Name "ExchangeOnline" -AllowBasicAuthActiveSync:$false -AllowBasicAuthImap:$false -AllowBasicAuthPop:$false      

Remove-AuthenticationPolicy -Identity "ExchangeOnlineDeskless"


$users = Get-User -ResultSize unlimited 
$id = $users.MicrosoftOnlineServicesID
$id | foreach {Set-User -Identity "chris hannah" -AuthenticationPolicy "MFAT Authentication Policy"}


Get-AuthenticationPolicy | Format-Table name -Auto

$SalesUsers = Get-User -ResultSize unlimited -Filter "(RecipientType -eq 'UserMailbox')"
$Sales = $SalesUsers.MicrosoftOnlineServicesID
$Sales | foreach {Set-User -Identity $_ -AuthenticationPolicy "ExchangeOnline"}




Set-CASMailboxPlan -Identity ExchangeOnline -ActiveSyncEnabled $False -ImapEnabled $false -PopEnabled $false

Get-CASMailboxPlan -Identity ExchangeOnlineDeskless | fl *enabled


Get-Organizationc