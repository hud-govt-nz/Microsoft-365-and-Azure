
#Import the Policies.json file
$Path = "C:\temp\Policies.json"
$json = Get-Content $Path | ConvertFrom-Json

#The script below generates the code required to create the teams policies

param(
    [Parameter(Mandatory=$true)]  
    [ValidateSet('TeamsChannel','Meeting','Messaging','Events')]
    [string]$PolicyType
)

$PolicyPrefix = "Restricted"
[string]$command = $null
 
switch ($PolicyType) {
    "TeamsChannel" {  
        $object = $json.Policies.Policy.Meeting
        $command += "New-CsTeamsChannelsPolicy"
        foreach($property in $object.psobject.properties)
        {
            if($property.name -eq "Identity")
            {
                $command += " -$($property.name) '$($PolicyPrefix) $($property.value)'"
            }
            else
            {
                $command += " -$($property.name) $($property.value)"
            }
        }
    }
    "Meeting" {  
        $object = $json.Policies.Policy.Messaging
        $command += "New-CsTeamsMeetingPolicy"
        foreach($property in $object.psobject.properties)
        {
            if($property.name -eq "Identity")
            {
                $command += " -$($property.name) '$($PolicyPrefix) $($property.value)'"
            }
            else
            {
                $command += " -$($property.name) $($property.value)"
            }
        }
    }
        "Messaging" {  
        $object = $json.Policies.Policy.Messaging
        $command += "New-CsTeamsMessagingPolicy"
        foreach($property in $object.psobject.properties)
        {
            if($property.name -eq "Identity")
            {
                $command += " -$($property.name) '$($PolicyPrefix) $($property.value)'"
            }
            else
            {
                $command += " -$($property.name) $($property.value)"
            }
        }
    }
        "Events" {  
        $object = $json.Policies.Policy.Messaging
        $command += "New-CsTeamsMeetingBroadcastPolicy"
        foreach($property in $object.psobject.properties)
        {
            if($property.name -eq "Identity")
            {
                $command += " -$($property.name) '$($PolicyPrefix) $($property.value)'"
            }
            else
            {
                $command += " -$($property.name) $($property.value)"
            }
        }
    }
}
return $command