$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2.0"
$validationFile = "$validation\HUD - New Device Setup.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}