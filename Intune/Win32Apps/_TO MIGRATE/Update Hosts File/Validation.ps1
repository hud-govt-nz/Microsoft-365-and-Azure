$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "1.1"
$validationFile = "$validation\Hosts_File_Update.txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
