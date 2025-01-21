## Set Variables:  
$members = New-Object System.Collections.ArrayList  
  
## Create the Function  
    function getMembership($group) {  
        $searchGroup = Get-DistributionGroupMember $group -ResultSize Unlimited  
        foreach ($member in $searchGroup) {  
            if ($member.RecipientTypeDetails-match "Group" -and $member.DisplayName -ne "") {  
                getMembership($member.DisplayName)  
                }             
            else {  
                if ($member.DisplayName -ne "") {  
                    if (! $members.Contains($member.DisplayName) ) {  
                        $members.Add($member.DisplayName) >$null  
                        }  
                    }  
                }  
            }  
        }  
   
## Run the function  
  
$groups = Get-DistributionGroup -Identity "DL - SDP - All Staff"
  
foreach ($group in $groups) {  
    write-host "`nGroup Name: " $group -ForegroundColor:Green  
    getMembership($group.DisplayName)  
    ## Output results to screen  
    write-host "Group Members:" -ForegroundColor:Yellow  
    $members.GetEnumerator() | sort-object  
    $members = New-Object System.Collections.ArrayList  
}