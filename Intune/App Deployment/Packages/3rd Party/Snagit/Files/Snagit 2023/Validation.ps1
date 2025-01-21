$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "23.1.1"
$validationFile = "$validation\Snagit 2023.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}