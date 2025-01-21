#Outline Model Types
$ApplicableModels = @(
    @{Model = 'Surface Pro 8'; Installer = 'C:\setup\SurfacePro8_Win11_22000_23.032.18675.0.msi'},
    @{Model = 'Surface Laptop 3'; Installer = 'c:\setup\SurfaceLaptop3_Win11_22000_23.021.14365.0.msi'},
    @{Model = 'Surface Laptop 4'; Installer = 'C:\setup\SurfaceLaptop4_Win11_22000_22.111.10120.0.msi'},
    @{Model = 'Surface Laptop 5'; Installer = 'C:\Setup\SurfaceLaptop5_Win11_22621_22.102.17243.0.msi'}
)

#Query Model of current device
$DeviceModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model 

#Install Respective Driver Pack
foreach ($model in $ApplicableModels)
{
    if ($DeviceModel -eq $model.Model) {
        Write-Host "The model of this device is $($DeviceModel). Running MSI DriverPack Installer..."
        Start-Process -FilePath $model.Installer -ArgumentList "/qb /norestart" -Wait 
        $ApplicableModels | Where-Object { $_.Model -ne $model.Model } | ForEach-Object {Remove-Item -Path $_.Installer -Force -Confirm:$false -ErrorAction SilentlyContinue}
        break
    }
}
