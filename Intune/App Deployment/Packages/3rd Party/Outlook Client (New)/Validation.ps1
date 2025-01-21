$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "17.10.35027.167"
$validationFile = "$validation\Visual Studio Community 2022.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}