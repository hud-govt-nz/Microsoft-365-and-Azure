$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2021.4.7"
$validationFile = "$validation\Snagit 2021.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}