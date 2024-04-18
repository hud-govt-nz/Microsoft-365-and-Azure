<#
.SYNOPSIS
This script installs the 22H2 Search Function Fix, version 1.0, and adds a validation marker in the "HUD Tools" folder within $env:homedrive\HUD Tools directory. 
It also creates log files in the "Log" folder within the "EndpointManager" folder.

.DESCRIPTION
The script first checks if HUD Font Install is already installed on the computer and exits the script if it is. 
It then checks if the specified directories and files already exist before trying to create them. 
The script then downloads and installs the HUD Font Install software, writes the version number to the validation file, and then deletes the temporary files.

.EXAMPLE
powershell .\22H2SearchFix.ps1 install
powershell .\22H2SearchFix.ps1 uninstall

.NOTES
- Make sure that you have the appropriate permissions to install software on the computer
- This script is designed to work with version 1.0 of HUD Font Install. If you need to install a different version, you will need to modify the script accordingly.
- This script assumes that the "HUD Tools" and "EndpointManager" folders already exist in the $env:homedrive\HUD directory. 
  If they do not exist, you will need to modify the script accordingly.

.AUTHOR
Ashley Forde

#>
############################################

#Test if root directory exists on device
$Folder = "$($env:homedrive)\HUD"

    if (Test-Path -Path $Folder) {"Path exists!"} else {
        "Creating root folder..."
        New-Item -Path "C:\" -Name "HUD" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "01_Logs" -ItemType "directory" -Force -Confirm:$false
        New-Item -Path "C:\HUD\" -Name "02_Validation" -ItemType "directory" -Force -Confirm:$false
    }

#Define working folders
$path = "$Folder\00_Staging\"
$logs = "$Folder\01_Logs\"
$validation = "$Folder\02_Validation\"

#Define application version
$version = "1.1"

#Define log file
$logfile = "$logs\22H2SearchFix.log"

############################################

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'action' {
        $Keys = Get-ChildItem -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Input\Locales
        $Keys | ForEach-Object {
            if ($_.Name -eq "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Input\Locales\Loc_0409" -or $_.Name -eq "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Input\Locales\Loc_0481" -or $_.Name -eq "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Input\Locales\Loc_0804") {
                "$($_.Name) exists on device"
                
                # Log the result of the installation
                Add-Content -Path $logfile -Value "[$(Get-Date)] $($_.Name) exists on device"

            } else {
                Remove-Item -Path "Registry::$($_.Name)" -Recurse -Force -Confirm:$false

                # Log the result of the installation
                Add-Content -Path $logfile -Value "[$(Get-Date)] $($_.Name) has been removed from device"

            }
        }
   
        # Create validation file
        New-Item -ItemType File -Path "$validation\22H2SearchFix.txt" -Force -Value $version
        Add-Content -Path $logfile -Value "[$(Get-Date)] Validation file was created successfully."

    }
}
