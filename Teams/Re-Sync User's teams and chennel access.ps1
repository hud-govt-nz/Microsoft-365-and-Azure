#Connect-MicrosoftTeams

#Obtain User
$User = Read-Host "Enter user UPN"

#Obtain Users Assigned Teams
$Teams = Get-Team -User $User

Foreach ($id in $Teams) {
    
    #Remove User from Teams first
    Remove-TeamUser  -GroupId $($id.GroupId) -User $User
   
    #Add user to team
    Write-Host "Adding $($User) to Group $($Id.DisplayName)..."
    Add-TeamUser -GroupId $($id.GroupId) -User $User -Role Member

}