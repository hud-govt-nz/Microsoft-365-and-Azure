#Connect to Teams 
Connect-MicrosoftTeams

#Export List of users in Teams Group
$Teams = Get-Team -GroupID "<enter group id here>"
$FolderPath = "$env:userprofile\Desktop\uniqueusers.csv"
$users = @()

#Loop through and export list of users in Teams Group
ForEach($i in $Teams.GroupId) {
    $users += Get-TeamUser -GroupId $i}
$uniqUsers = $users | sort UserId -Unique
$uniqUsers | Export-Csv -Path $FolderPath 

#Filter/Check names on list...


#Import Filtered list and remove users from Teams Group 
$Import = Import-csv 'C:\Users\xo-aforde\Desktop\Filter Team Org Wide Group List.csv'
$GroupID = 'd0239d16-efb7-49f8-bfd1-e638158ddec2'

foreach ($User in $Import) {
    Remove-TeamUser -GroupId $GroupID -User $User.userid}