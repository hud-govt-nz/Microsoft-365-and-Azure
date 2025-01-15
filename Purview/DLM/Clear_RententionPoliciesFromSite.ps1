#Parameter
$SiteURL = "https://crescent.sharepoint.com/sites/Marketing"
$PolicyName = "Retention Policy"
 
Try {
    #Connect to Compliance Center through Exchange Online Module
    Connect-IPPSSession
 
    #Get the Policy
    $Policy = Get-RetentionCompliancePolicy -Identity $PolicyName
 
    If($Policy)
    {
        #Exclude site from Retention Policy
        Set-RetentionCompliancePolicy -AddSharePointLocationException $SiteURL -Identity $Policy.Name
    }
}
Catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}


#Read more: https://www.sharepointdiary.com/2021/07/how-to-exclude-sharepoint-online-site-from-retention-policy.html#ixzz8UDi87drQ