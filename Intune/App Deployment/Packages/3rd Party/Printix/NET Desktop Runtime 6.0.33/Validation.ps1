$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "6.0"
$validationFile = "$validation\Microsoft Windows Desktop Runtime - 6.0.33.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}