$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "1.0"
$validationFile = "$validation\Suppress Retention Policy in Outlook (Classic) UI.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
