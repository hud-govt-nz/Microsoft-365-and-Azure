# Specify the group name  
$groupName = "iOS-Corp-Mgmt-Devices"  
$group = Get-MgGroup -Filter "displayName eq '$groupName'"  
$groupId = $group.Id

# Get members of the group  
$Devices = Get-MgGroupMember -GroupId $groupId
$DeviceIds = @()

# Get the device IDs of the registered owners of the devices
Foreach ($Device in $Devices) {
    $registeredOwners = (Get-MgDevice -DeviceId $Device.id -ExpandProperty "registeredOwners").RegisteredOwners
    foreach ($owner in $registeredOwners) {
        $DeviceIds += $owner.Id
    }
}

# Add the device IDs as members of the group
foreach ($deviceId in $DeviceIds) {
    $group = "64336224-1ec2-4146-9c3d-dea55a17be94"
    New-MgGroupMember -GroupId $group -DirectoryObjectId $deviceId
}