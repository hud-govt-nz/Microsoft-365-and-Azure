# Define output file and remove if the same file already exists in the directory path
$currentDir = $(Get-Location).Path
$oFile = "$($currentDir)\AzureADUsersAndRoleAssignments.csv"

if(Test-Path -Path $oFile){Remove-Item $oFile -Force}

# Login to Azure AD
Connect-AzureAD

$allAZADUserWithRoleMapping = @()

# Get all Azure AD roles and loop through members of those roles
# Add user/service principal details in psObject array
Get-AzureADDirectoryRoleTemplate | ForEach-Object{
    $roleName = $_.DisplayName
    Get-AzureADDirectoryRole | Where-Object {$_.displayName -eq $roleName} | ForEach-Object{
        Get-AzureADDirectoryRoleMember -ObjectId $_.ObjectId | ForEach-Object{
            $extProp = $_.ExtensionProperty
            $objUser = New-Object psObject
            $objUser | Add-Member RoleName  $roleName
            $objUser | Add-Member UserName $_.DisplayName
            $objUser | Add-Member JobTitle $_.JobTitle
            $objUser | Add-Member EMail $_.Mail
            $objUser | Add-Member AccountEnabled $_.AccountEnabled
            $objUser | Add-Member Department $_.Department
            $objUser | Add-Member ObjectType $_.ObjectType
            $objUser | Add-Member CreationDate $extProp.createdDateTime
            $objUser | Add-Member EmployeeId  $extProp.employeeId  
            $allAZADUserWithRoleMapping += $objUser
        }
    }
}
$allAZADUserWithRoleMapping | Export-CSV -Path $oFile -NoClobber -NoTypeInformation -Confirm:$false -Force