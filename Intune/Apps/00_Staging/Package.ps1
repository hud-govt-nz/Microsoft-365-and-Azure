[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$AppName = Read-Host -Prompt "Enter the folder name for the package"

$PackageFolder = New-Object -Typename System.Windows.Forms.FolderBrowserDialog
$PackageFolder.rootfolder = "Desktop"
$PackageFolder.ShowDialog()

$FolderName = "$($packageFolder.SelectedPath)\$AppName"

if (-not (Test-Path $packageFolder)) {
    New-Item -Path $FolderName -ItemType "directory" -Force
}
else {
    Write-Host "Folder $FolderName already exists. Moving on."
}

Copy-Item -Path "$PSScriptRoot\functions.ps1" -Destination $FolderName -Force

$DeploymentFolder = New-Object -Typename System.Windows.Forms.FolderBrowserDialog
$DeploymentFolder.rootfolder = "Desktop"
$DeploymentFolder.ShowDialog()


Copy-Item -Path "$($DeploymentFolder.SelectedPath)\*" -Destination $FolderName -Recurse -Force

IntuneWinAppUtil.exe -c "$folderName" -s "appinstall.ps1" -o "C:\HUD\20_Packages" -q