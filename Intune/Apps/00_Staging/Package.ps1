[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

Write-Host "Please enter the name for this package" -ForegroundColor Yellow
$AppName = Read-Host -Prompt "Enter the folder name for the package"


Write-Host "Please select the folder for the package" -ForegroundColor Yellow

$PackageFolder = New-Object -Typename System.Windows.Forms.FolderBrowserDialog
$PackageFolder.SelectedPath = "C:\HUD\20_Packages"
$PackageFolder.ShowDialog()


$FolderName = "$($packageFolder.SelectedPath)\$AppName"
Write-Host "Folder name: $FolderName" -ForegroundColor Green

if (-not (Test-Path $packageFolder)) {
    New-Item -Path $FolderName -ItemType "directory" -Force
}
else {
    Write-Host "Folder $FolderName already exists. Moving on." -ForegroundColor Green
}

Write-Host "Copying common functions to $FolderName" -ForegroundColor Green
Copy-Item -Path "$PSScriptRoot\functions.ps1" -Destination $FolderName -Force

Write-host "Please select the folder for the deployment" -ForegroundColor Yellow
$DeploymentFolder = New-Object -Typename System.Windows.Forms.FolderBrowserDialog

$DeploymentFolder.SelectedPath = "$env:appfolder"
$DeploymentFolder.ShowDialog()


Copy-Item -Path "$($DeploymentFolder.SelectedPath)\*" -Destination $FolderName -Recurse -Force
Write-host "Files copied to $FolderName" -ForegroundColor Green

Write-host "Starting IntuneWinAppUtil.exe and compiling .intunewin file" -ForegroundColor Cyan
IntuneWinAppUtil.exe -c "$folderName" -s "appinstall.ps1" -o "C:\HUD\20_Packages" -q 