$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "22.200"
$validationFile = "$validation\Oracle Smart View 64-bit for Office.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}