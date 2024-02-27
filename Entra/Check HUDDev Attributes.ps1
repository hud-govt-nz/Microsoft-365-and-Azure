Connect-MgGraph -Scopes "Directory.Read.All","Directory.ReadWrite.All","User.Read.All","User.ReadWrite.All" -TenantId 1432e95a-a79e-4c3c-ac12-ec29543a8286

$User = Get-MgBetaUser -UserId "Andrew.Bazzi@huddev.onmicrosoft.com"

Write-Host "User $($User.DisplayName) has the following extension attributes:" -ForegroundColor Green
Write-Host "Leave Date Time: $($User.AdditionalProperties.extension_2babf13c01ce40b1a19d0b489a4d53f2_ObjectUserLeaveDateTime)"
Write-Host "Employee Type: $($User.AdditionalProperties.extension_2babf13c01ce40b1a19d0b489a4d53f2_ObjectUserEmployeeType)"
Write-Host "Organisational Group: $($User.AdditionalProperties.extension_2babf13c01ce40b1a19d0b489a4d53f2_ObjectUserOrganisationalGroup)"
Write-Host "Employment Category: $($User.AdditionalProperties.extension_2babf13c01ce40b1a19d0b489a4d53f2_ObjectUserEmploymentCategory)"
Write-Host "Start Date: $($User.AdditionalProperties.extension_2babf13c01ce40b1a19d0b489a4d53f2_ObjectUserStartDate)"
