# Script parameters and configuration
[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = "C:\HUD",
    [switch]$IncludeTeamsPolicy
)

# Check for ImportExcel module and install if missing
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Import required module
Import-Module ImportExcel

# Initialize logging
$LogPath = Join-Path $OutputPath "ComplianceReport_$(Get-Date -Format 'yyyy-MM-dd_HH-mm').log"
function Write-Log {
    param($Message)
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogPath -Value $LogMessage
}

Write-Log "Starting compliance policy report generation..."

try {
    # Create timestamp and export path for Excel file
    $Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $ExcelPath = Join-Path $OutputPath "RetentionPolicy_Report_$Timestamp.xlsx"

    # Initialize data arrays for each sheet
    $PolicyData = [System.Collections.ArrayList]::new()
    $RulesData = [System.Collections.ArrayList]::new()
    $LocationsData = [System.Collections.ArrayList]::new()

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
        $PercentComplete = ($CurrentPolicyIndex / $TotalPolicies) * 100
        
        Write-Progress -Id 1 -Activity "Processing Retention Compliance Policies" `
            -Status "Policy $CurrentPolicyIndex of $TotalPolicies : $($Policy.Name)" `
            -PercentComplete $PercentComplete `
            -CurrentOperation "Initializing policy processing..."
        
        Write-Log "Processing policy: $($Policy.Name)"
        
        try {
            # Add policy data
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing Policy Details" `
                -Status "Adding policy information" `
                -PercentComplete 25

            $PolicyData.Add([PSCustomObject]@{
                PolicyName = $Policy.Name
                PolicyType = $Policy.Type
                PolicyEnabled = $Policy.Enabled
                PolicyMode = $Policy.Mode
                PolicyWorkload = $Policy.Workload
                LastModified = $Policy.LastModifiedTime
                CreatedTime = $Policy.WhenCreated
            }) | Out-Null

            # Get and process rules
            $Rules = Get-RetentionComplianceRule -Policy $Policy.Guid
            $TotalRules = $Rules.Count
            Write-Log "Found $TotalRules rules for policy '$($Policy.Name)'"
            
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing Policy Rules" `
                -Status "Processing $TotalRules rules" `
                -PercentComplete 50
            
            $CurrentRuleIndex = 0
            foreach ($Rule in $Rules) {
                $CurrentRuleIndex++
                $RulePercent = ($CurrentRuleIndex / $TotalRules) * 100
                
                Write-Progress -Id 3 -ParentId 2 -Activity "Processing Rules" `
                    -Status "Rule $CurrentRuleIndex of $TotalRules" `
                    -PercentComplete $RulePercent `
                    -CurrentOperation $Rule.Name

                $RulesData.Add([PSCustomObject]@{
                    PolicyName = $Policy.Name
                    RuleName = $Rule.Name
                    ComplianceTagProperty = $Rule.ComplianceTagProperty
                    PublishComplianceTag = $Rule.PublishComplianceTag
                    RuleMode = $Rule.Mode
                    RuleWorkload = $Rule.Workload
                }) | Out-Null
            }
            
            # Process SharePoint locations
            Write-Progress -Id 2 -ParentId 1 -Activity "Processing Locations" `
                -Status "Processing SharePoint locations" `
                -PercentComplete 75
            
            if ($null -ne $Policy.SharePointLocation) {
                if ($Policy.SharePointLocation.Name -eq "All") {
                    $LocationsData.Add([PSCustomObject]@{
                        PolicyName = $Policy.Name
                        Type = "SharePointLocation"
                        Name = "All SharePoint Sites"
                        URL = "All SharePoint Sites"
                    }) | Out-Null
                } else {
                    $TotalLocations = $Policy.SharePointLocation.Count
                    $CurrentLocationIndex = 0
                    
                    foreach ($location in $Policy.SharePointLocation) {
                        $CurrentLocationIndex++
                        $LocationPercent = ($CurrentLocationIndex / $TotalLocations) * 100
                        
                        Write-Progress -Id 3 -ParentId 2 -Activity "Processing SharePoint Locations" `
                            -Status "Location $CurrentLocationIndex of $TotalLocations" `
                            -PercentComplete $LocationPercent `
                            -CurrentOperation $location.DisplayName
                        
                        $LocationsData.Add([PSCustomObject]@{
                            PolicyName = $Policy.Name
                            Type = "SharePointLocation"
                            Name = $location.DisplayName
                            URL = $location.Name
                        }) | Out-Null
                    }
                }
            }

            # Process SharePoint location exceptions
            if ($null -ne $Policy.SharePointLocationException) {
                $TotalExceptions = $Policy.SharePointLocationException.Count
                $CurrentExceptionIndex = 0
                
                foreach ($location in $Policy.SharePointLocationException) {
                    $CurrentExceptionIndex++
                    $ExceptionPercent = ($CurrentExceptionIndex / $TotalExceptions) * 100
                    
                    Write-Progress -Id 3 -ParentId 2 -Activity "Processing Location Exceptions" `
                        -Status "Exception $CurrentExceptionIndex of $TotalExceptions" `
                        -PercentComplete $ExceptionPercent `
                        -CurrentOperation $location.DisplayName
                    
                    $LocationsData.Add([PSCustomObject]@{
                        PolicyName = $Policy.Name
                        Type = "SharePointLocationException"
                        Name = $location.DisplayName
                        URL = $location.Name
                    }) | Out-Null
                }
            }
        }
        catch {
            Write-Log "Error processing policy '$($Policy.Name)': $_"
            continue
        }
        
        # Clear child progress bars
        Write-Progress -Id 3 -Completed
        Write-Progress -Id 2 -Completed
    }

    Write-Progress -Id 1 -Activity "Finalizing Report" -Status "Creating Excel file..." -PercentComplete 99

    # Export to Excel using Export-Excel cmdlet
    $ExcelParams = @{
        Path = $ExcelPath
        AutoSize = $true
        AutoFilter = $true
        FreezeTopRow = $true
        BoldTopRow = $true
        TableStyle = 'Medium2'
        WorksheetName = 'Policies'
    }

    # Export Policies
    $PolicyData | Export-Excel @ExcelParams

    # Export Rules
    $RulesData | Export-Excel -Path $ExcelPath -WorksheetName 'Rules' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle 'Medium2'

    # Export Locations
    $LocationsData | Export-Excel -Path $ExcelPath -WorksheetName 'Locations' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle 'Medium2'

    Write-Progress -Id 1 -Completed
    
    Write-Log "Excel report has been generated successfully:"
    Write-Log "Report location: $ExcelPath"
}
catch {
    Write-Log "Critical error during report generation: $_"
    throw
}
finally {
    # Clear any remaining progress bars
    Write-Progress -Id 3 -Completed
    Write-Progress -Id 2 -Completed
    Write-Progress -Id 1 -Completed
    
    Write-Log "Script execution completed"
}