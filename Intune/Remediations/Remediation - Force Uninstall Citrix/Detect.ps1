$citrixComponents = Get-WmiObject -Class Win32_Product | 
    Where-Object {$_.Name -match "citrix"}

if ($citrixComponents) {
    Write-Host "Citrix components found. Remediation needed."
    Exit 1
} else {
    Write-Host "No Citrix components found."
    Exit 0
}