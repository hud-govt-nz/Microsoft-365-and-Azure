Connect-MgGraph -NoWelcome

$commands =@('Get-MGUser')

# Initialize an array to collect all permissions
$allPermissions = @()

# Iterate over each command and collect the permissions
foreach ($Command in $Commands) {
    $permissions = Find-MgGraphCommand -Command $Command | Select-Object -First 1 -ExpandProperty permissions
    $allPermissions += $permissions
}


# Initialize a hashtable to keep track of unique names
$uniqueNames = @{}

# Iterate over all permissions and add unique names to the hashtable
foreach ($permission in $allPermissions) {
    $name = $permission.Name  # Assuming the Name property holds the name of the permission
    if (-not $uniqueNames.ContainsKey($name)) {
        $uniqueNames[$name] = $permission
    }
}

# sort and output the unique permissions
$sortedUniquePermissions = $uniqueNames.Values | Sort-Object -Property Name
$sortedUniquePermissions.name