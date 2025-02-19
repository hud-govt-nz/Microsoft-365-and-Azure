$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "<APP VERSION>"
$validationFile = "$validation\<APP NAME>.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}