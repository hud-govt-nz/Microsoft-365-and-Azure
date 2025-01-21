$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "11.1.14"
$validationFile = "$validation\Postman.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}