$TargetScope = "/subscriptions/0fc3e508-28eb-427a-b4fa-f78bdc39f2c6/resourceGroups/OD4PolicyTestRG"
$resourceTypes = @()
$path = 
$csv = import-csv $path
foreach ($resource in $csv.List) {$resourceTypes += $Resource}

Definition and Assignment creation:

$definition = /subscriptions/0fc3e508-28eb-427a-b4fa-f78bdc39f2c6/providers/Microsoft.Authorization/policyDefinitions/bba2f932-0b62-48cb-a321-d3d5c6eb2acc
$assignment = New-AzPolicyAssignment -Name 'testing-allowed-resources' -Scope $TargetScope -listOfResourceTypesAllowed $resourceTypes -PolicyDefinition $definition