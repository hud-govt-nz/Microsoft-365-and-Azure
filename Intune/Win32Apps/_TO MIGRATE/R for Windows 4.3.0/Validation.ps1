$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "4.3.0"
$validationFile = "$validation\R for Windows 4.3.0.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}