$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "7.4.5"
$validationFile = "$validation\PowerShell 7-x64.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}