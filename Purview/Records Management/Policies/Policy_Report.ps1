# Retrieve all retention compliance policies
$policies = Get-RetentionCompliancePolicy

# Initialize an array to collect report results
$report = @()

$policyCount = $policies.Count
$currentPolicyIndex = 0

foreach ($policy in $policies) {
    $currentPolicyIndex++
    Write-Progress -Id 1 -Activity "Processing Retention Compliance Policies" -Status "Processing policy: $($policy.Name)" -PercentComplete (($currentPolicyIndex / $policyCount) * 100)
    
    # Get all rules associated with the policy (using its Guid)
    $rules = Get-RetentionComplianceRule -Policy $policy.Guid
    
    $ruleCount = $rules.Count
    $currentRuleIndex = 0
    
    foreach ($rule in $rules) {
        $currentRuleIndex++
        Write-Progress -Id 2 -ParentId 1 -Activity "Processing Rules" -Status "Processing rule: $($rule.Name)" -PercentComplete (($currentRuleIndex / $ruleCount) * 100)
        
        $report += [PSCustomObject]@{
            PolicyName            = $policy.Name
            ComplianceTagProperty = $rule.ComplianceTagProperty
            PublishComplianceTag  = $rule.PublishComplianceTag
            Policy                = $rule.Policy
            ObjectVersion         = $rule.ObjectVersion
            Guid                  = $rule.Guid
            Id                    = $rule.Id
        }
    }
}

Write-Progress -Id 1 -Activity "Processing Retention Compliance Policies" -Completed
Write-Progress -Id 2 -Activity "Processing Rules" -Completed

# Display the collected results in a table format
$report | Export-csv -Path "C:\HUD\RetentionCompliancePolicyReport.csv" -NoTypeInformation