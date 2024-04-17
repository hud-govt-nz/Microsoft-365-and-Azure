<#
.SYNOPSIS
This script installs the HUD Support Tool, version 1.0, in the "HUD Tools" folder within $env:homedrive\HUD Tools directory. 
It also creates log files in the "Log" folder within the "EndpointManager" folder.

.DESCRIPTION
The script first checks if HUD Support Tools is already installed on the computer and exits the script if it is. 
It then checks if the specified directories and files already exist before trying to create them. 
The script then downloads and installs the HUD Support Tools software, writes the version number to the validation file, and then deletes the temporary files.

.EXAMPLE
powershell .\HUD Support Tool.ps1 install
powershell .\HUD Support Tool.ps1 uninstall

.NOTES
- Make sure that you have the appropriate permissions to install software on the computer
- This script is designed to work with version 1.0 of HUD Support Tools. If you need to install a different version, you will need to modify the script accordingly.
- This script assumes that the "HUD Tools" and "EndpointManager" folders already exist in the $env:homedrive\HUD Tools directory. 
  If they do not exist, you will need to modify the script accordingly.

.AUTHOR
Ashley Forde

#>
############################################

# Test if root directory exists on device
$Folder = "$($env:homedrive)\HUD"

    if (Test-Path -Path $Folder) {
        "Path exists!"
    } else {
        "Creating root folder..."
        New-Item -Path "C:\" -Name "HUD" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "01_Logs" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "02_Validation" -ItemType "directory" -Force -Confirm:$false
    }

# Define working folders
$path = "$Folder\00_Staging\"
$logs = "$Folder\01_Logs\"
$validation = "$Folder\02_Validation\"

# Define application version
$version = "1.0"

# Define log file
$logfile = "$logs\HUD Support Tool.log"

############################################

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        #  Copy HUD Support Tools Icon files to local device
        $ScriptFile = New-Item -Path $Folder -Name 05_SupportTool -ItemType Directory -Force -Confirm:$false
        Copy-Item -Path ".\Installer\*" -Destination $ScriptFile -Recurse -Force
        
        $scriptPath = "$ScriptFile\InvokeTool.ps1"
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $shortcutPath = Join-Path $desktopPath "Digital Support Admin Shell.lnk"
        
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = "powershell.exe"
        $Shortcut.Arguments = "-ExecutionPolicy Bypass -File `"$scriptPath`""
        $Shortcut.Save()


        # Create validation file
        New-Item -ItemType File -Path "$validation\HUD Support Tool.txt" -Force -Value $version

        # Log that validation file was created
        Add-Content -Path $logfile -Value "[$(Get-Date)] Validation file was created successfully."
       
    }
}
#Refresh Desktop
$wsh = New-Object -ComObject Wscript.Shell
$wsh.sendkeys('{F5}')
exit