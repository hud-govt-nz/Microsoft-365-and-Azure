<# Set Teams Policies via AAD Group Membership

    .PURPOSE
     The function is setup to directly assign MS Teams policies to AAD users based on whether or not they belong to a specific security group in AAD.

    .PROCESS
     1. Define paramters GroupName & relevant policy 
     2. Connect to Azure AD & MS Teams Powershell Modules
     3. Apply the relevant Policy directly to all the users whom are members of the Security Group.

    .NOTE
     Applying policies in this manner shows as a direct assignment within the Teams Admin Centre, that means it will supersede any policies assigned to users via AAD policy assignment. 
     Further information can be found here:
     https://docs.microsoft.com/en-us/microsoftteams/assign-policies-users-and-groups#:~:text=Policy%20assignment%20to%20groups%20lets%20you%20assign%20a,group%2C%20their%20inherited%20policy%20assignments%20are%20updated%20accordingly.

    Created by: Ashley Forde
    Version: 1
    Date: 15.4.22

#>

#Connect to Azure MSOL and Teams
Connect-MsolService
Connect-MicrosoftTeams

#Function
#Meeting Policy
function Set-TeamsMeetingPolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsMeetingPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams meeting policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#Messaging Policy
function Set-TeamsMessagingPolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsMessagingPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams messaging policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#Live Events Policy
function Set-TeamsLiveEventsPolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsMeetingBroadcastPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams live events policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#App Permissions
function Set-TeamsAppPermissionPolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsAppPermissionPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams app permission policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#App Setup Policy
function Set-TeamsAppSetupPolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsAppSetupPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams app setup policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#Teams (Channel) Policy
function Set-TeamsChannelPolicy {  
    param ($GroupName,$PolicyName) 
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsChannelsPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams channel policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}
#Upgrade Policy
function Set-TeamsUpgradePolicy {  
    param ($GroupName,$PolicyName)
        process{
            #Get security group information.
            $group= Get-MsolGroup -SearchString $GroupName |select ObjectId,DisplayName
            $members=Get-MsolGroupMember -GroupObjectId $group.ObjectId -MemberObjectTypes user -all

            #Add user to App permission policy
            foreach($member in $members){
                Grant-CsTeamsUpdateManagementPolicy -PolicyName $PolicyName -Identity $member.EmailAddress
                Write-Host "Teams upgrade policy successfully added to $($member.EmailAddress) user " 
                } 
        }
}


<# EXAMPLES
Set-TeamsMeetingPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted - Meeting Policy"
Set-TeamsMessagingPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted - Messaging Policy"
Set-TeamsLiveEventsPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted - Live Events Policy"
Set-TeamsAppPermissionPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted - App Permission Policy"
Set-TeamsAppSetupPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted - App Setup Policy"
Set-TeamsChannelPolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted  - Teams Channel Policy"
Set-TeamsUpgradePolicy -GroupName "AAD-Teams-Restricted-User" -PolicyName "Restricted User - App Update Policy"


Set-TeamsMeetingPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Meeting Policy"
Set-TeamsMessagingPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Messaging"
Set-TeamsLiveEventsPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Live Events"
Set-TeamsAppPermissionPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - App Permissions"
Set-TeamsAppSetupPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Setup"
Set-TeamsChannelPolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Teams Policy"
Set-TeamsUpgradePolicy -GroupName "AAD-Teams-CloudOperations" -PolicyName "Cloud Operations - Update Policy"
#>