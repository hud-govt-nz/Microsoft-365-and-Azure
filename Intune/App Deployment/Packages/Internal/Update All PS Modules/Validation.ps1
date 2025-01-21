$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "0.1"
$validationFile = "$validation\PowerShell Module Update.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}