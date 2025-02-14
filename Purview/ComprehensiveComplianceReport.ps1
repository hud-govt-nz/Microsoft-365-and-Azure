# Script parameters and configuration
[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = "C:\HUD",
    [switch]$IncludeTeamsPolicy
)

# Initialize logging
$LogPath = Join-Path $OutputPath "ComplianceReport_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').log"
function Write-Log {
    param($Message)
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogPath -Value $LogMessage
}

# Initialize lists to store the reports
$ComprehensiveReport = [System.Collections.Generic.List[Object]]::new()
$DetailedRulesReport = [System.Collections.Generic.List[Object]]::new()

# Helper function to add policy location report lines
function Add-PolicyLocationReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$PolicyName,
        [Parameter(Mandatory)]
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
    
    try {
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
            ReportGenerated = Get-Date
        }
        $ComprehensiveReport.Add($ReportLine)
    }
    catch {
        Write-Log "Error adding report line for policy '$PolicyName': $_"
    }
}

Write-Log "Starting compliance policy report generation..."

try {
    # Get retention policies based on parameters
    $PolicyParams = @{
        DistributionDetail = $true
    }
    if (-not $IncludeTeamsPolicy) {
        $PolicyParams.Add('ExcludeTeamsPolicy', $true)
    }
    
    Write-Log "Retrieving retention compliance policies..."
    $Policies = Get-RetentionCompliancePolicy @PolicyParams
    Write-Log "Found $($Policies.Count) policies to process"
    
    $TotalPolicies = $Policies.Count
    $CurrentPolicyIndex = 0

    # Process each policy
    foreach ($Policy in $Policies) {
        $CurrentPolicyIndex++
        Write-Progress -Id 1 -Activity "Processing Retention Compliance Policies" `
            -Status "Processing policy: $($Policy.Name)" `
            -PercentComplete (($CurrentPolicyIndex / $TotalPolicies) * 100)
        
        Write-Log "Processing policy: $($Policy.Name)"
        
        try {
            # Get all rules associated with the policy
            $Rules = Get-RetentionComplianceRule -Policy $Policy.Guid
            Write-Log "Found $($Rules.Count) rules for policy '$($Policy.Name)'"
            
            # Process rules for detailed report
            foreach ($Rule in $Rules) {
                try {
                    $DetailedRulesReport.Add([PSCustomObject]@{
                        PolicyName = $Policy.Name
                        RuleName = $Rule.Name
                        ComplianceTagProperty = $Rule.ComplianceTagProperty
                        PublishComplianceTag = $Rule.PublishComplianceTag
                        Policy = $Rule.Policy
                        ObjectVersion = $Rule.ObjectVersion
                        Guid = $Rule.Guid
                        Id = $Rule.Id
                        ReportGenerated = Get-Date
                    })
                }
                catch {
                    Write-Log "Error processing rule '$($Rule.Name)' for policy '$($Policy.Name)': $_"
                    continue
                }
            }

            # Process SharePoint locations
            if ($null -ne $Policy.SharePointLocation) {
                if ($Policy.SharePointLocation.Name -eq "All") {
                    Add-PolicyLocationReport -PolicyName $Policy.Name -PolicyType $Policy.Type `
                        -LocationType "SharePointLocation" -SiteName "All SharePoint Sites" `
                        -SiteURL "All SharePoint Sites" -Workload $Policy.Workload `
                        -Enabled $Policy.Enabled -Mode $Policy.Mode `
                        -ComplianceTags ($Rules.ComplianceTagProperty -join "; ") `
                        -PublishTags ($Rules.PublishComplianceTag -join "; ") `
                        -LabelWorkload ($Rules.Workload -join "; ") `
                        -LabelPolicy ($Rules.Policy -join "; ") -LabelMode ($Rules.Mode -join "; ")
                } else {
                    $Policy.SharePointLocation | ForEach-Object {
                        Add-PolicyLocationReport -PolicyName $Policy.Name -PolicyType $Policy.Type `
                            -LocationType "SharePointLocation" -SiteName $_.DisplayName `
                            -SiteURL $_.Name -Workload $Policy.Workload `
                            -Enabled $Policy.Enabled -Mode $Policy.Mode `
                            -ComplianceTags ($Rules.ComplianceTagProperty -join "; ") `
                            -PublishTags ($Rules.PublishComplianceTag -join "; ") `
                            -LabelWorkload ($Rules.Workload -join "; ") `
                            -LabelPolicy ($Rules.Policy -join "; ") -LabelMode ($Rules.Mode -join "; ")
                    }
                }
            }

            # Process SharePoint location exceptions
            if ($null -ne $Policy.SharePointLocationException) {
                $Policy.SharePointLocationException | ForEach-Object {
                    Add-PolicyLocationReport -PolicyName $Policy.Name -PolicyType $Policy.Type `
                        -LocationType "SharePointLocationException" -SiteName $_.DisplayName `
                        -SiteURL $_.Name -Workload $Policy.Workload `
                        -Enabled $Policy.Enabled -Mode $Policy.Mode `
                        -ComplianceTags ($Rules.ComplianceTagProperty -join "; ") `
                        -PublishTags ($Rules.PublishComplianceTag -join "; ") `
                        -LabelWorkload ($Rules.Workload -join "; ") `
                        -LabelPolicy ($Rules.Policy -join "; ") -LabelMode ($Rules.Mode -join "; ")
                }
            }
        }
        catch {
            Write-Log "Error processing policy '$($Policy.Name)': $_"
            continue
        }
    }

    Write-Progress -Id 1 -Activity "Processing Retention Compliance Policies" -Completed
    
    # Create timestamp and export paths
    $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $DetailedRulesPath = Join-Path $OutputPath "RetentionPolicy_DetailedRules_$Timestamp.csv"
    $LocationsReportPath = Join-Path $OutputPath "RetentionPolicy_Locations_$Timestamp.csv"

    # Export the reports
    $DetailedRulesReport | Export-Csv -Path $DetailedRulesPath -NoTypeInformation
    $ComprehensiveReport | Export-Csv -Path $LocationsReportPath -NoTypeInformation

    Write-Log "Reports have been generated successfully:"
    Write-Log "1. Detailed Rules Report: $DetailedRulesPath"
    Write-Log "2. Policy Locations Report: $LocationsReportPath"
}
catch {
    Write-Log "Critical error during report generation: $_"
    throw
}
finally {
    Write-Log "Script execution completed"
}