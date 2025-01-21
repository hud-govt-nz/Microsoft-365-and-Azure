$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "125.0.6422.142"
$validationFile = "$validation\Google Chrome Enterprise.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}