$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "4.12"
$validationFile = "$validation\Citrix Receiver 4.12.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
