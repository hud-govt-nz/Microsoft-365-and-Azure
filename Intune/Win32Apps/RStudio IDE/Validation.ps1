$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2023.03.0+386"
$validationFile = "$validation\RStudio.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}