Set-Location C:\HUD\06_Reporting
  
# Authenticate with the Graph  
Connect-MgGraph -NoWelcome
#Select-MgProfile -Name beta  

# Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Get all groups from Microsoft Graph
$allGroups = Get-MGGroup -All | Where-Object { $_.GroupTypes -notcontains "Unified" -and $_.DisplayName -notlike "DMS*" -and $_.DisplayName -notcontains "DL - *" -and $_.DisplayName -notlike "PIM - *"}

# Get all Distribution Groups from Exchange Online
$allDistGroups = Get-DistributionGroup | Select-Object -ExpandProperty DisplayName

# Filter out distribution groups from $allGroups
$allGroups = $allGroups | Where-Object { $allDistGroups -notcontains $_.DisplayName }
  
# Function to get nested groups  
function Get-NestedGroups($rootGroupId, $groupId, $layer, $nestedGroupNames, $groupTypes) {  
    $nestedGroups = Get-MGGroupMemberOf -GroupId $groupId  
    $output = @()  
  
    foreach ($group in $nestedGroups) {  
        $nestedGroupName = $group.AdditionalProperties["displayName"]  
        $nestedGroupNames[$layer] = $nestedGroupName  
        $obj = [PSCustomObject]@{  
            "id" = $rootGroupId  
            "Group" = $nestedGroupNames[0]  
            "GroupTypes" = ($groupTypes -join ', ')  
            "Nested Layer 1" = $nestedGroupNames[1]  
            "Nested Layer 2" = $nestedGroupNames[2]  
            "Nested Layer 3" = $nestedGroupNames[3]  
            "Nested Layer 4" = $nestedGroupNames[4]  
        }  
        $output += $obj  
  
        if ($layer -lt 5) {  
            $output += Get-NestedGroups -rootGroupId $rootGroupId -GroupId $group.Id -layer ($layer + 1) -nestedGroupNames $nestedGroupNames -groupTypes $group.GroupTypes  
        }  
    }  
    return $output  
}  
  
# Display all groups and their nested groups  
$results = @()  
$groupCount = $allGroups.Count  
  
foreach ($group in $allGroups) {  
    $nestedGroupNames = @($group.DisplayName, "", "", "", "", "")  
    $groupResult = Get-NestedGroups -rootGroupId $group.Id -GroupId $group.Id -layer 1 -nestedGroupNames $nestedGroupNames -groupTypes $group.GroupTypes  
  
    if ($groupResult.Count -eq 0) {  
        $obj = [PSCustomObject]@{  
            "id" = $group.Id  
            "Group" = $group.DisplayName  
            "GroupTypes" = ($group.GroupTypes -join ', ')  
            "Nested Layer 1" = ""  
            "Nested Layer 2" = ""  
            "Nested Layer 3" = ""  
            "Nested Layer 4" = ""  
        }  
        $results += $obj  
    } else {  
        $results += $groupResult  
    }  
  
    $progress = [math]::Round((($results.Count) / $groupCount) * 100, 2)  
    $progress = [math]::Min($progress, 100)  
    Write-Progress -Activity "Processing Groups" -Status "Processing $($group.DisplayName)" -PercentComplete $progress  
}  
  
Write-Progress -Activity "Processing Groups" -Status "Completed" -Completed  
  
# Now, we have all the results stored in $results  
$Date = (Get-Date).AddDays(0).ToString("dd-MMM-yy") 
$results | Export-Excel -Path ".\Nested_Entra_Groups.xlsx" -Append -WorksheetName "$Date" -FreezeTopRow -AutoSize -AutoFilter
