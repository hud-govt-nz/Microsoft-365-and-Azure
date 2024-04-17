Clear-Host 

# Fetch all data first to reduce the number of API calls per group
$AllApps = Get-IntuneMobileApp -Select id, displayName, lastModifiedDateTime, assignments -Expand assignments
$AllDeviceCompliance = Get-IntuneDeviceCompliancePolicy -Select id, displayName, lastModifiedDateTime, assignments -Expand assignments
$AllDeviceConfig = Get-IntuneDeviceConfigurationPolicy -Select id, displayName, lastModifiedDateTime, assignments -Expand assignments

$Resource = "deviceManagement/deviceManagementScripts"
$graphApiVersion = "Beta"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$expand=groupAssignments"
$AllDeviceConfigScripts = Invoke-MSGraphRequest -HttpMethod GET -Url $uri

$Resource = "deviceManagement/groupPolicyConfigurations"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$expand=Assignments"
$AllADMT = Invoke-MSGraphRequest -HttpMethod GET -Url $uri

$Groups = Get-MGGroup -All | Where-Object { $_.GroupTypes -notcontains "Unified" -and $_.DisplayName -notlike "DMS*" -and $_.DisplayName -notlike "DL - *" -and $_.DisplayName -notlike "PIM - *" }

$results = @()

$totalGroups = $Groups.Count
$currentGroupNumber = 0

Foreach ($Group in $Groups) {
    $currentGroupNumber++
    Write-Progress -PercentComplete (($currentGroupNumber / $totalGroups) * 100) -Status "Processing Group $currentGroupNumber of $totalGroups" -Activity "Scanning Groups"

    $groupResult = [PSCustomObject]@{
        'AAD Group Name'                  = $Group.displayName
        'Apps'                            = $AllApps | Where-Object {$_.assignments -match $Group.id} | ForEach-Object { $_.DisplayName }
        'Compliance'                      = $AllDeviceCompliance | Where-Object {$_.assignments -match $Group.id} | ForEach-Object { $_.DisplayName }
        'Configurations'                  = $AllDeviceConfig | Where-Object {$_.assignments -match $Group.id} | ForEach-Object { $_.DisplayName }
        'PowershellScripts'               = $AllDeviceConfigScripts.value | Where-Object {$_.assignments -match $Group.id} | ForEach-Object { $_.DisplayName }
        'AdministrativeTemplates'         = $AllADMT.value | Where-Object {$_.assignments -match $Group.id} | ForEach-Object { $_.DisplayName }
    }

    # Output positive results
    if ($groupResult.'Apps'.Count -gt 0) {
        Write-Host ""
        Write-Host "Group $($Group.displayName) is assigned to the following apps: $($groupResult.'Apps' -join ', ')" -ForegroundColor Green
    }

    if ($groupResult.'Device Compliance Policies'.Count -gt 0) {
        Write-Host ""
        Write-Host "Group $($Group.displayName) is assigned to the following Device Compliance Policies: $($groupResult.'Device Compliance Policies' -join ', ')" -ForegroundColor Green
    }

    if ($groupResult.'Device Configurations'.Count -gt 0) {
        Write-Host ""
        Write-Host "Group $($Group.displayName) is assigned to the following Device Configurations: $($groupResult.'Device Configurations' -join ', ')" -ForegroundColor Green
    }

    if ($groupResult.'Device Config Powershell Scripts'.Count -gt 0) {
        Write-Host ""
        Write-Host "Group $($Group.displayName) is assigned to the following Device Config Powershell Scripts: $($groupResult.'Device Config Powershell Scripts' -join ', ')" -ForegroundColor Green
    }

    if ($groupResult.'Administrative Templates'.Count -gt 0) {
        Write-Host ""
        Write-Host "Group $($Group.displayName) is assigned to the following Administrative Templates: $($groupResult.'Administrative Templates' -join ', ')" -ForegroundColor Green
    }

    $results += $groupResult
}

$modifiedResults = $results | ForEach-Object {
    [PSCustomObject]@{
        'AAD Group'                       = $_.'AAD Group Name'
        'Apps'                            = ($_.Apps -join ', ')
        'Device Compliance Policies'      = ($_.Compliance -join ', ')
        'Device Configurations'           = ($_.Configurations -join ', ')
        'Device Config Powershell Scripts'= ($_.PowershellScripts -join ', ')
        'Administrative Templates'        = ($_.AdministrativeTemplates -join ', ')
    }
}

$modifiedResults | Export-Csv -Path "C:\HUD\IntuneAssignmentReport.csv" -NoTypeInformation

