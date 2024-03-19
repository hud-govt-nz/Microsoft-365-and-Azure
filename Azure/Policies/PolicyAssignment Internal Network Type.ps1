$TargetScope = "/subscriptions/0fc3e508-28eb-427a-b4fa-f78bdc39f2c6/ResourceGroups/Od4T2CorRl01"
$resourceTypes = @()
$path = "C:\PowerShell\Azure\Policies\Blacklist-AzPolicy.csv"
$csv = import-csv $path
foreach ($resource in $csv.List) {$resourceTypes += $Resource}

$definition = Get-AzPolicyDefinition -name 6c112d4e-5bc7-47ae-a041-ea2d9dccd749 
$assignment = New-AzPolicyAssignment -Name 'OD4_Od4T2CorRl01_BlockedList' -Scope $TargetScope -listOfResourceTypesAllowed $resourceTypes -PolicyDefinition $definition
