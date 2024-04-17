$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "22000_23.042.26034.0"
$validationFile = "$validation\Surface Laptop 4 Driver Pack.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}