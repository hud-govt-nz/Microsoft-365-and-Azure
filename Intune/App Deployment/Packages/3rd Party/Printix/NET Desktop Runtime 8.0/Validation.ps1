$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "8.0.11.34221"
$validationFile = "$validation\Microsoft Windows Desktop Runtime - 8.0.11.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}