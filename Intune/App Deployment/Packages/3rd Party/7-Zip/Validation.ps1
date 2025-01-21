$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "24.07"
$validationFile = "$validation\7-Zip.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
