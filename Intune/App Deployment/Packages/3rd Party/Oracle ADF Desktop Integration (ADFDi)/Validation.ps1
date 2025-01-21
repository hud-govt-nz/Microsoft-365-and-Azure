$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "5.2.0.26990"
$validationFile = "$validation\Oracle ADF Desktop Integration Add-In for Excel.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}