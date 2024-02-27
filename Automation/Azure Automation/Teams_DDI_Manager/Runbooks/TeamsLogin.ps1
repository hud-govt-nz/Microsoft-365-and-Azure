#Runbook to login as a system-assigned managed identity
$login = Get-AutomationPSCredential -Name 'TeamsAdminAccount'
Connect-MicrosoftTeams -Credential $login
