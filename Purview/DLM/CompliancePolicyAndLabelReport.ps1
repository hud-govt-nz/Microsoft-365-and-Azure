# Initialize a list to store the report
$Report = [System.Collections.Generic.List[Object]]::new()

# Helper function to add report lines
function Add-ReportLine {
    param (
        [string]$PolicyName,
        [string]$PolicyType,
        [string]$LocationType,
        [string]$SiteName,
        [string]$SiteURL,
        [string]$Workload,
        [bool]$Enabled,
        [string]$Mode,
        [string]$ComplianceTags,
        [string]$PublishTags,
        [string]$LabelWorkload,
        [string]$LabelPolicy,
        [string]$LabelMode
    )
    $ReportLine = [PSCustomObject]@{
        PolicyName = $PolicyName
        PolicyType = $PolicyType
        LocationType = $LocationType
        SiteName = $SiteName
        SiteURL = $SiteURL
        Workload = $Workload
        Enabled = $Enabled
        Mode = $Mode
        ComplianceTags = $ComplianceTags
        PublishTags = $PublishTags
        LabelWorkload = $LabelWorkload
        LabelPolicy = $LabelPolicy
        LabelMode = $LabelMode
    }
    $Report.Add($ReportLine)
}

# Retrieve the first 10 retention compliance policies
$Policies = Get-RetentionCompliancePolicy -DistributionDetail -ExcludeTeamsPolicy
$TotalPolicies = $Policies.Count
$CurrentPolicyIndex = 0

# Iterate through each policy
ForEach ($Policy in $Policies) {
    $CurrentPolicyIndex++
    Write-Progress -Activity "Processing Policies" -Status "Processing $CurrentPolicyIndex of $TotalPolicies" -PercentComplete (($CurrentPolicyIndex / $TotalPolicies) * 100)

    # Retrieve the retention compliance rules associated with the policy
    $LabelDetails = Get-RetentionComplianceRule -Policy $Policy.Guid

    # Check if SharePointLocation is not null
    if ($null -ne $Policy.SharePointLocation) {
        # Check if the policy applies to all SharePoint sites
        if ($Policy.SharePointLocation.Name -eq "All") {
            Add-ReportLine -PolicyName $Policy.Name -PolicyType $Policy.Type -LocationType "SharePointLocation" `
                -SiteName "All SharePoint Sites" -SiteURL "All SharePoint Sites" -Workload $Policy.Workload `
                -Enabled $Policy.Enabled -Mode $Policy.Mode -ComplianceTags $LabelDetails.ComplianceTagProperty `
                -PublishTags $LabelDetails.PublishComplianceTag -LabelWorkload $LabelDetails.Workload `
                -LabelPolicy $LabelDetails.Policy -LabelMode $LabelDetails.Mode
        } else {
            # Expand SharePoint locations
            $Policy.SharePointLocation | ForEach-Object {
                Add-ReportLine -PolicyName $Policy.Name -PolicyType $Policy.Type -LocationType "SharePointLocation" `
                    -SiteName $_.DisplayName -SiteURL $_.Name -Workload $Policy.Workload `
                    -Enabled $Policy.Enabled -Mode $Policy.Mode -ComplianceTags $LabelDetails.ComplianceTagProperty `
                    -PublishTags $LabelDetails.PublishComplianceTag -LabelWorkload $LabelDetails.Workload `
                    -LabelPolicy $LabelDetails.Policy -LabelMode $LabelDetails.Mode
            }
        }
    }

    # Check if SharePointLocationException is not null
    if ($null -ne $Policy.SharePointLocationException) {
        # Expand SharePoint location exceptions
        $Policy.SharePointLocationException | ForEach-Object {
            Add-ReportLine -PolicyName $Policy.Name -PolicyType $Policy.Type -LocationType "SharePointLocationException" `
                -SiteName $_.DisplayName -SiteURL $_.Name -Workload $Policy.Workload `
                -Enabled $Policy.Enabled -Mode $Policy.Mode -ComplianceTags $LabelDetails.ComplianceTagProperty `
                -PublishTags $LabelDetails.PublishComplianceTag -LabelWorkload $LabelDetails.Workload `
                -LabelPolicy $LabelDetails.Policy -LabelMode $LabelDetails.Mode
        }
    }
}

# Export the unique report to a CSV file
$Report | Export-Csv -Path "C:\HUD\CompliancePolicyandLabelReport.csv" -NoTypeInformation