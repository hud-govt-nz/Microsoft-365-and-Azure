function Get-UniqueGraphPermissions {
    <#
    .SYNOPSIS
    Retrieves unique Microsoft Graph permissions for a given command.

    .PARAMETER command
    The command for which to retrieve permissions, specified as a string.

    .RETURN
    A table containing the unique permissions with the header "Scopes".
    #>
    param(
        [string]$command
    )

    Connect-MgGraph -NoWelcome

    [array]$allPermissions = @()

    [array]$permissions = Find-MgGraphCommand -Command $command | Select-Object -First 1 -ExpandProperty permissions
    $allPermissions += $permissions

    [hashtable]$uniqueNames = @{}

    foreach ($permission in $allPermissions) {
        $name = $permission.Name
        if (-not $uniqueNames.ContainsKey($name)) {
            $uniqueNames[$name] = $permission
        }
    }

    [array]$sortedUniquePermissions = $uniqueNames.Values | Sort-Object -Property Name

    $table = $sortedUniquePermissions | Select-Object -Property @{Name="Scopes"; Expression={$_.Name}}
    Write-Output $table
}