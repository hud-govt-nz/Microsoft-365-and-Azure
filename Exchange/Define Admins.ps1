Connect-ExchangeOnline

$FormatEnumerationLimit=-1
$Roles = Get-RoleGroup | Select-Object Name, Members
$Result = Foreach ($Role in $Roles) {
    $Name = $Role.Name
    $Members = $Role | Select-Object -ExpandProperty Members
    [pscustomobject]@{
        'Name' = [string]$Name
        'Members' = [string]$Members
    }
}
$Result | Format-Table -Wrap 
$Result | Export-csv -Path C:\Support\Exportechangeadmin.csv -NoTypeInformation -Delimiter ';'



#Update Permissions for Users
$RoleIdentity = Read-Host 'enter role identity here i.e. help desk, recipient administrators'
$Users =@('Francis Amarasingha','Norman Niro')

if($RoleIdentity){
    Write-Host "Updating role membership for exhange role group $RoleIdentity"
        foreach ($User in $Users) {
            
            Update-RoleGroupMember -Identity $RoleIdentity -Members $User -Confirm:$false
}
}
