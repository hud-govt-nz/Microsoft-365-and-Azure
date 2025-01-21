$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "6.2.7.49583"
$validationFile = "$validation\Zoom Workplace x64.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}