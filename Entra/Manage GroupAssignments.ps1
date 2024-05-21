
$Scopes =@("DeviceManagementManagedDevices.Read.All", "DeviceManagementManagedDevices.ReadWrite.All","Directory.Read.All", "Group.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All","Directory.Read.All","Organization.Read.All","User.Read","User.Read.All","User.ReadWrite.All")
Connect-MgGraph -Scopes $Scopes

# Variables
$iOSPersonalMgmtUsers = "2e11dff9-e35a-48eb-b3c9-6677b8e6da30"
$iOSCorpMgmtUsers = "64336224-1ec2-4146-9c3d-dea55a17be94"

# Data Collection
$CorpActiveUsers =@()

$Corp_Device_User      = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'iOS' and managedDeviceOwnerType eq 'Company'" | select id,deviceName,managedDeviceOwnerType,operatingSystem,userPrincipalName,userId
$Corp_Group_User       = Get-MgGroupMember -GroupId $iOSPersonalMgmtUsers | select id


$Corp_Device_User | ForEach-Object {
    
    $User.displayName = Get-MgBetaUser -UserId $_.UserId | select id,accountEnabled,displayName
    $CorpActiveUsers += $User
}


      
$Personal_Device_User  = Get-MgDeviceManagementManagedDevice -Filter "operatingSystem eq 'iOS' and managedDeviceOwnerType eq 'Personal'" | select id,deviceName,managedDeviceOwnerType,operatingSystem,userPrincipalName,userId
$Personal_Group_User   = Get-MgGroupMember -GroupId $iOSCorpMgmtUsers | select id,userPrincipalName,accountEnabled