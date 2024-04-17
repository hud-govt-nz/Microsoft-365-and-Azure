$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "22000_23.041.9917.0"
$validationFile = "$validation\Surface Pro 8 Driver Pack.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}