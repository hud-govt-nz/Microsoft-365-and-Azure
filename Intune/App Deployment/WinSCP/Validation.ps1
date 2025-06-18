$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "6.3.5"
$validationFile = "$validation\WinSCP.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}