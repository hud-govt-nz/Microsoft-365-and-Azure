# Import the Microsoft Graph module
#Import-Module -Name Microsoft.Graph

# Function to get the group Id
function Get-GroupId($displayName)
{
    $group = Get-MgGroup -Filter "displayName eq '$displayName'"
    return $group.Id
}

# Function to delete the group
function Delete-Group($groupId)
{
    Remove-MgGroup -GroupId $groupId -Confirm:$false
}

# Authenticate to Graph
Connect-MgGraph

# The display name of the group you want to remove
$groupName = 'your-group-name-here'

# Get the group Id
$groupId = Get-GroupId -displayName $groupName

# Delete the group
Delete-Group -groupId $groupId

# Disconnect from Graph
Disconnect-MgGraph
