$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "3.0"
$validationFile = "$validation\HUD Email Signature.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
