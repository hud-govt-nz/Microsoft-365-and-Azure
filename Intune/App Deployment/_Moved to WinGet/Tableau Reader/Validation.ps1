$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "24.2.931.0"
$validationFile = "$validation\Tableau Reader.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}