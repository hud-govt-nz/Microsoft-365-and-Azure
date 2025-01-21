$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "4.3.3"
$validationFile = "$validation\R for Windows.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
