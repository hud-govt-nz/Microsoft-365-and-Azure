$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "3.1"
$validationFile = "$validation\HUD - Office 365 Document Templates.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
