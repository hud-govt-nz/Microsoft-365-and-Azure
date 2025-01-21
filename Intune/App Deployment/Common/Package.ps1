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

Write-host "Please select the folder of the deployment files" -ForegroundColor Yellow
$DeploymentFolder = New-Object -Typename System.Windows.Forms.FolderBrowserDialog

$DeploymentFolder.SelectedPath = "C:\HUD\15_CodeRepo\AshForde\M365-Staging\INTUNE\App Deployment\Packages"
$DeploymentFolder.ShowDialog()


Copy-Item -Path "$($DeploymentFolder.SelectedPath)\*" -Destination $FolderName -Recurse -Force
Write-host "Files copied to $FolderName" -ForegroundColor Green

Write-Host "Please copy the installation files into the package folder: $FolderName" -ForegroundColor Yellow
Write-Host "Press any key to continue after copying the files..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Write-host "Starting IntuneWinAppUtil.exe and compiling .intunewin file" -ForegroundColor Cyan
& C:\HUD\10_Software\Microsoft-Win32-Content-Prep-Tool-master\IntuneWinAppUtil.exe -c "$folderName" -s "appinstall.ps1" -o "C:\HUD\20_Packages" -q 

Write-Host "Opening Explorer at package location" -ForegroundColor Cyan
Invoke-Item -Path $FolderName