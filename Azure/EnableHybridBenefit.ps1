#Sign into Azure - Will sign in using azure dev admin account
Connect-AzAccount
 
#Get Windows Only VM's under Subscription - check output before updating licenseType
$WinVM = Get-AzVM -Status | where{$_.StorageProfile.OsDisk.OsType -eq "Windows" -and (!($_.LicenseType))} | Export-Csv -Path C:\Temp\WinVM.csv
 
#Import WinVM list
$list = Import-Csv -Path C:\temp\WinVM.csv
 
#Enable Hybrid Benefit / Update License
foreach($i in $list)
{
    $vm = Get-AzVM -ResourceGroup $($i.ResourceGroupName) -Name $($i.Name)
    Write-Verbose "Setting hybrid benefit on VM $($i.Name) "
    $vm.LicenseType = "Windows_Server"
    Update-AzVM -ResourceGroupName $($i.ResourceGroupName) -VM $vm
}
 
#Verify
Get-AzVM -Status | select ResourceGroupName, Name, LicenseType | out-file C:\temp\Export.txt