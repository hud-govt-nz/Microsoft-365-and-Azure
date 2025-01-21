$Result = @()
#Get all teams
$AllTeams= Get-Team
$TotalTeams = $AllTeams.Count
$i = 0
#Iterate teams one by one and get channels 
ForEach ($Team in $AllTeams)
{
$i++
Write-Progress -Activity "Fetching channels from $($Team.Displayname)" -Status "$i out of $TotalTeams completed"
Try
{
#Get channels
$TeamChannels = Get-TeamChannel -GroupId $Team.GroupId
 
#Iterate channels one by one and add to the result array
ForEach ($Channel in $TeamChannels)
{
#Add channel info to the result array
$Result += New-Object PSObject -property $([ordered]@{
TeamName = $Team.DisplayName
TeamVisibility = $Team.Visibility
ChannelName = $Channel.DisplayName
GroupId = $Team.GroupId
ChannelId = $Team.ChannelId
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
$Result | Export-CSV "C:\Powershell\Teams\AllTeamChannels.CSV" -NoTypeInformation -Encoding UTF8