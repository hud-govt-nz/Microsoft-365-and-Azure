$Folder = "$($env:homedrive)\HUD"
$validation = "$Folder\02_Validation"
$version = "1.0"
$validationFile = "$validation\Te Reo Maori Greetings Outlook Plugin.txt"
$content = Get-Content -Path $validationFile

if($content -eq $version){
    Write-Host "Found it!"
}