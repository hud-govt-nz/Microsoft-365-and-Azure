$TargetScope = "/subscriptions/0fc3e508-28eb-427a-b4fa-f78bdc39f2c6"
$resourceTypes = @()
$path = "C:\PowerShell\Azure\Policies\Whitelist-AzPolicy.csv"
$csv = import-csv $path
foreach ($resource in $csv.List) {$resourceTypes += $Resource}

$definition = Get-AzPolicyDefinition -name bba2f932-0b62-48cb-a321-d3d5c6eb2acc 
$assignment = New-AzPolicyAssignment -Name 'OD4_Subscription_AllowedList' -Scope $TargetScope -listOfResourceTypesAllowed $resourceTypes -PolicyDefinition $definition
