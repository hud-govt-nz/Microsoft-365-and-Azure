$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2.0"
$validationFile = "$validation\HUD - Disable Outlook Auto Complete.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}