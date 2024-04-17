$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "5.1.1.24107"
$validationFile = "$validation\Oracle ADF Desktop Integration Add-In for Excel.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}