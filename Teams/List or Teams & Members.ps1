$AllTeamsInOrg = (Get-Team).GroupID
$TeamList = @()

Write-Output "This may take a little bit of time... Please sit back, relax and enjoy some GIFs inside of Teams!"
Write-Host ""

Foreach ($Team in $AllTeamsInOrg)
{       
        $TeamGUID = $Team.ToString()
        $TeamGroup = Get-UnifiedGroup -identity $Team.ToString()
        $TeamName = (Get-Team | ?{$_.GroupID -eq $Team}).DisplayName
        $TeamOwner = (Get-TeamUser -GroupId $Team | ?{$_.Role -eq 'Owner'}).User
        $TeamUserCount = ((Get-TeamUser -GroupId $Team).UserID).Count
        $TeamGuest = (Get-UnifiedGroupLinks -LinkType Members -identity $Team | ?{$_.Name -match "#EXT#"}).Name
            if ($TeamGuest -eq $null)
            {
                $TeamGuest = "No Guests in Team"
            }
        $TeamChannels = (Get-TeamChannel -GroupId $Team).DisplayName
	$ChannelCount = (Get-TeamChannel -GroupId $Team).ID.Count
        $TeamList = $TeamList + [PSCustomObject]@{TeamName = $TeamName; TeamObjectID = $TeamGUID; TeamOwners = $TeamOwner -join ', '; TeamMemberCount = $TeamUserCount; NoOfChannels = $ChannelCount; ChannelNames = $TeamChannels -join ', '; SharePointSite = $TeamGroup.SharePointSiteURL; AccessType = $TeamGroup.AccessType; TeamGuests = $TeamGuest -join ','}
}

#######

$TestPath = test-path -path 'c:\temp'
if ($TestPath -ne $true) {New-Item -ItemType directory -Path 'c:\temp' | Out-Null
    write-Host  'Creating directory to write file to c:\temp. Your file is uploaded as TeamsDatav2.csv'}
else {Write-Host "Your file has been uploaded to c:\temp as 'TeamsDatav2.csv'"}
$TeamList | export-csv c:\temp\TeamsDatav2.csv -NoTypeInformation