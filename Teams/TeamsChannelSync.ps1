#Logging
$Date = Get-date -Format "dd.MM.yyyy"
Start-Transcript C:\Support\Teams_Channel_Sync\Transcript\Teams_Channel_ReSync_$Date.txt -Append

#Step 1: Connect to Azure and MicrosoftTeams
#Connect-AzureAD
#Connect-MicrosoftTeams
<#
#Step 2: Obtain list of licensed users and add to array
$Result = @()
Get-AzureAdUser -All $true | ForEach-Object { 
    $licensed=$False ; For ($i=0; $i -le ($_.AssignedLicenses | Measure-Object).Count ; $i++) { 
        If( [string]::IsNullOrEmpty(  $_.AssignedLicenses[$i].SkuId ) -ne $True) {
            $licensed=$true 
            }
        } ; If( $licensed -eq $true) {
                $result += New-Object psobject -Property $([ordered]@{UPN = $_.UserPrincipalName})
                }
    }
#Export list of licensed users in tenant for record keeping
$Result | Export-CSV "C:\Support\Teams_Channel_Sync\UPNList\UPN.CSV" -NoTypeInformation -Encoding UTF8 -Force
#>

$Result = Import-csv -Path "C:\Support\Teams_Channel_Sync\UPNList\AAD_ObjectID__Export_202211160734.CSV"
Foreach ($user in $result) {

$Team =""
$Teams =@()
$Channel =""
$Channels =@()

#Get Teams
Get-Team -User $User.userPrincipalName | ForEach-Object {
    $TeamName = $_.DisplayName
    $GroupId= $_.GroupId
        
    if (([string]$TeamName -match [string]'AIP Users') -or ([string]$TeamName -match [string]'Te Pae Kōrero') ) { 
        Write-host "Skipping $($TeamName)" -ForegroundColor red
    } else {
        #Get Team Users
        Get-TeamUser -GroupId $GroupId | ForEach-Object {
            $Name= $_.Name
            $MemberMail= $_.User
            $Role= $_.Role
            $OwnerCount = (Get-TeamUser -GroupId $GroupId -Role Owner).count

            #Save Value if user is in Team
            if ([string]$MemberMail -match [string]$User.userPrincipalName) {
                Write-Host "$($Name) is in Team: $($TeamName) with: $($Role) permissions" -ForegroundColor Yellow
                $Team =@{'TeamName'=$TeamName; 'Name'=$Name; 'UPN'=$MemberMail; 'Role'=$Role}
                $Teams = New-Object psobject -Property $Team
                $Teams | Select-Object 'TeamName', 'Name', 'UPN', 'Role' | Export-Csv -Path C:\Support\Teams_Channel_Sync\TeamsUserList\TeamsUserList_$($User.userPrincipalName)_$Date.csv -NoTypeInformation -Append -Force

                    if (([string]$Team.Role -match 'owner') -and ($OwnerCount -gt 1)) {
                        #Remove Permissions
                        Write-Host "Resetting $($Role) permissions on Group $($TeamName) for $($MemberMail)..." -ForegroundColor Cyan
                        Remove-TeamUser -GroupId $GroupId -User $MemberMail -Role $Role
                        #Add Permissions
                        Add-TeamUser -GroupId $GroupId -User $MemberMail -Role $Role
                        Write-Host "$($Role) permissions on Group $($TeamName) for $($MemberMail) have been reset." -ForegroundColor Green
                        } 
                        elseif (([string]$MemberMail -match [string]$User.userPrincipalName) -and ([string]$Team.Role -match 'owner') -and ($OwnerCount -eq 1)) {
                            Write-Host "$($User.userPrincipalName) is the sole owner of group: $($TeamName). Please refresh permissions manually" -ForegroundColor Red
                            }

                    if (([string]$MemberMail -match [string]$User.userPrincipalName) -and ([string]$Team.Role -eq 'member') -or ([string]$Team.Role -eq 'guest')) {
                        #Remove Permissions
                        Write-Host "Resetting $($Role) permissions on Group $($TeamName) for $($MemberMail)..." -ForegroundColor Cyan
                        Remove-TeamUser -GroupId $GroupId -User $MemberMail -Role $Role
                        #Add Permissions
                        Add-TeamUser -GroupId $GroupId -User $MemberMail -Role $Role
                        Write-Host "$($Role) permissions on Group $($TeamName) for $($MemberMail) have been reset." -ForegroundColor Green
                        }
                }
            }

        #Get Channels within each Team
        Get-TeamChannel -GroupId $GroupId | ForEach-Object {
            $channelDispName = $_.DisplayName
            $channelmbrType = $_.MembershipType 
            Write-Host "$($TeamName) has the following Channel: $($channelDispName)" -ForegroundColor DarkYellow

            #Get Channel Users
            Get-TeamChannelUser -GroupId $GroupId -DisplayName $channelDispName | ForEach-Object {
                $CName = $_.Name
                $CUser = $_.User
                $CRole = $_.Role
                
                #Save Value if user is in Channel
                if([string]$CUser -match [string]$User.userPrincipalName) {
                    Write-Host "$($CName) has the role: $($CRole) in the channel: $($channelDispName) in Team: '$($TeamName)'" -ForegroundColor DarkCyan
                    $Channel =@{'TeamName'=$TeamName; 'Channel'=$channelDispName; 'Name'=$CName; 'ChannelRole'=$CRole}
                    $Channels = New-Object psobject -Property $Channel
                    $Channels | Select-Object 'TeamName', 'Channel', 'Name', 'ChannelRole' | Export-Csv -Path C:\Support\Teams_Channel_Sync\ChannelUserList\ChannelUserList_$($User.userPrincipalName)_$Date.csv -NoTypeInformation -Append -Force
                    
                    if ([string]$channelmbrType -eq 'Private') {
                        #Remove Permissions
                        Write-Host "Updating Private Channel $($channelDispName) permissions for user $($CUser)" -ForegroundColor DarkBlue
                        Remove-TeamChannelUser -GroupId $GroupId -DisplayName $channelDispName -User $CUser -Role Owner
                        #Add Permissions
                        Add-TeamChannelUser -GroupId $GroupId -DisplayName $channelDispName -User $CUser -Role Owner
                        Write-Host "Permissions for Private Channel $($channelDispName) have been reset for user $($CUser)" -ForegroundColor DarkBlue
                        }
                    }
                }
            }
        } 
    }
    [GC]::Collect()
}
Stop-Transcript

