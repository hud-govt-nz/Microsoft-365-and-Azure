$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "17.6.4"
$validationFile = "$validation\Microsoft Visual Studio 2022 Community Edition.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}