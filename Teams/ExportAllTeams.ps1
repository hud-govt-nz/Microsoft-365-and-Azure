Connect-MicrosoftTeams

#List all the teams 
Get-Team
 
#List the private teams
Get-Team -Visibility Private
 
#List the archived teams
Get-Team -Archived $true

Get-Team | Export-CSV "C:\Powershell\Teams\AllTeams.CSV" -NoTypeInformation -Encoding UTF8