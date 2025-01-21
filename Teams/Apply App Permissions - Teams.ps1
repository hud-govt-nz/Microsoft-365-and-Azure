Connect-MicrosoftTeams


$Group = Get-Team -DisplayName "APEC Incident Management"
$Users = Get-TeamUser -GroupId $Group.Groupid -Role Member
$Users | ForEach-Object {Grant-CsTeamsAppPermissionPolicy -PolicyName "APEC Incident Management - App Permissions" -Identity $_.User}


$Group = Get-Team -DisplayName "Cloud Programme"
$Users = Get-TeamUser -GroupId $Group.Groupid -Role Member
$Users | ForEach-Object {Grant-CsTeamsAppPermissionPolicy -PolicyName "Cloud Operations - App Permissions" -Identity $_.User}