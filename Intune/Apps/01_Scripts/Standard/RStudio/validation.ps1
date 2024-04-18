$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "2023.12.1.0"
$validationFile = "$validation\RStudio for Windows.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
