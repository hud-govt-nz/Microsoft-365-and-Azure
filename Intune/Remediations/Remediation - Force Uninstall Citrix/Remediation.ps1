$citrixComponents = Get-WmiObject -Class Win32_Product | 
    Where-Object {$_.Name -match "citrix"}

$exitCode = 0

foreach ($component in $citrixComponents) {
    Write-Host "Uninstalling $($component.Name)..." -ForegroundColor Yellow
    try {
        $component.Uninstall()
        Write-Host "Successfully uninstalled $($component.Name)" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to uninstall $($component.Name): $($_.Exception.Message)" -ForegroundColor Red
        $exitCode = 1
    }
}

Exit $exitCode