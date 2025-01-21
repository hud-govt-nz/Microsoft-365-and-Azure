$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2025"
$validationFile = "$validation\Printix Desktop Client.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}