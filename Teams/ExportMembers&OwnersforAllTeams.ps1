$Result = @()
#Get all teams
$AllTeams= Get-Team
$TotalTeams = $AllTeams.Count
$i = 0
#Iterate teams one by one and get channels 
ForEach ($Team in $AllTeams)
{
$i++
Write-Progress -Activity "Fetching users for $($Team.Displayname)" -Status "$i out of $TotalTeams completed"
Try
{
#Get team users
$TeamUsers = Get-TeamUser -GroupId $Team.GroupId
 
#Iterate users one by one and add to the result array
ForEach ($TeamUser in $TeamUsers)
{
#Add user info to the result array
$Result += New-Object PSObject -property $([ordered]@{
TeamName = $Team.DisplayName
TeamVisibility = $Team.Visibility
UserName = $TeamUser.Name
UserPrincipalName = $TeamUser.User
Role = $TeamUser.Role
GroupId = $Team.GroupId
UserId = $Team.UserId
})
}
}
Catch 
{
Write-Host "Error occurred for $($Team.Displayname)" -f Yellow
Write-Host $_ -f Red
}
}
 
#Export the result to CSV file
$Result | Export-CSV "C:\Powershell\Teams\AllTeamOwnersandMembers.CSV" -NoTypeInformation -Encoding UTF8