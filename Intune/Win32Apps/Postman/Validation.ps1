$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "10.14.2"
$validationFile = "$validation\Postman x86_64 10.14.2.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}