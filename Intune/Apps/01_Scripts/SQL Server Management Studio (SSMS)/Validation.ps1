$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "20.1"
$validationFile = "$validation\SQL Server Management Studio (SSMS).txt"
$content = Get-Content -Path $validationFile

if ($content -eq $version) {
	Write-Host "Found it!"
}
