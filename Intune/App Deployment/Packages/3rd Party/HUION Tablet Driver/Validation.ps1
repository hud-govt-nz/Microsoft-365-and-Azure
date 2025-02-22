$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "15.7.6.1314"
$validationFile = "$validation\HuionTablet_WinDriver.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}