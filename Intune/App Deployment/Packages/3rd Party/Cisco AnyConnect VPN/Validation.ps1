$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "5.1.7.80"
$validationFile = "$validation\Cisco Secure Client - AnyConnect VPN.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}