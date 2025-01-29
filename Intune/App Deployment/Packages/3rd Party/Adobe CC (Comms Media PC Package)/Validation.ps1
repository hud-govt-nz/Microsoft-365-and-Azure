$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "<APPVERSION>"
$validationFile = "$validation\Adobe Acrobat Reader.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}