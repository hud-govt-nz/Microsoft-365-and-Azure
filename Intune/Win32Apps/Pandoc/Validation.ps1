$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "3.1.2"
$validationFile = "$validation\Pandoc 3.1.2.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}