$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "10.24"
$validationFile = "$validation\Postman Desktop App v10.24.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
