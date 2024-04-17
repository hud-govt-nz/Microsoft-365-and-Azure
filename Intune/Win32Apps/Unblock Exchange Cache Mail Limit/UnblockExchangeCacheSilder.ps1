<#
.SYNOPSIS
This script installs the Outlook Remove Exchange Cache Silder Policy script, version 1.0, in the "HUD Tools" folder within $env:homedrive\HUD Tools directory. 
It also creates log files in the "Log" folder within the "EndpointManager" folder.

.DESCRIPTION
The script first checks if Outlook Remove Exchange Cache Silder Policy script is already installed on the computer and exits the script if it is. 
It then checks if the specified directories and files already exist before trying to create them. 
The script then downloads and installs the Outlook Remove Exchange Cache Silder Policy script, writes the version number to the validation file, and then deletes the temporary files.

.EXAMPLE
powershell .\Outlook Remove Exchange Cache Silder Policy script.ps1 install
powershell .\Outlook Remove Exchange Cache Silder Policy script.ps1 uninstall

.NOTES
- Make sure that you have the appropriate permissions to run script on the computer
- This script is designed to work with version 1.0 of Outlook Remove Exchange Cache Silder Policy script. If you need to install a different version, you will need to modify the script accordingly.
- This script assumes that the "HUD Tools" and "EndpointManager" folders already exist in the $env:homedrive\HUD Tools directory. 
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
$version = "1.0"

#Define log file
$logfile = "$logs\Outlook Remove Exchange Cache Silder Policy.log"

############################################

# Obtain Current User SID to enter into 
$currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1] 
$Keys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse
    Foreach($Key in $Keys) {
        if(($key.GetValueNames() | ForEach-Object{$key.GetValue($_)}) -match $CurrentUser ){$sid = $key}
        }
    # SID for current User
    $SID = $sid.pschildname 

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        # Remove legacy script items
        $User= (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
        
        # Old validation and logging files
        Get-Childitem -path "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Log" | where-object {$_.Name -ilike "*W11_Remove_EXO_CacheSliderPolicy-Install.log*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Validation"| where-object {$_.Name -ilike "*W11_Remove_EXO_CacheSliderPolicy*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Data"| where-object {$_.Name -ilike "*W11_Remove_EXO_CacheSliderPolicy*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Log"| where-object {$_.Name -ilike "*Remove_EXO_CacheModeSettings-Install*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Validation"| where-object {$_.Name -ilike "*Remove_EXO_CacheModeSettings*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Program Files\HUD Tools\EndpointManager\Log\" | where-object {$_.Name -ilike "*W11_Remove_EXO_CacheSliderPolicy-Install.log*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
        Get-Childitem -path "C:\Program Files\HUD Tools\EndpointManager\Validation\" | where-object {$_.Name -ilike "*W11_Remove_EXO_CacheSliderPolicy*"} | Remove-Item -Force -Confirm:$false -ErrorAction SilentlyContinue
             
        #Note these values are not dynamically removed by intune when a user is added to this policy, hence this script which manually removes them after they are added.   
        Remove-ItemProperty -Path "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\" -Name syncwindowsettingdays -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Remove-ItemProperty -Path "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\" -Name syncwindowsetting -Force -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
       
        #Restart Outlook
        Get-Process Outlook* | Stop-Process -Force -Confirm:$false -ErrorAction SilentlyContinue

        Add-Content -Path $logfile -Value "[$(Get-Date)] Outlook Remove Exchange Cache Silder Policy script version $($version) was installed successfully."

        # Create validation file
        New-Item -ItemType File -Path "$validation\Outlook Remove Exchange Cache Silder Policy.txt" -Force -Value $version

        # Log that validation file was created
        Add-Content -Path $logfile -Value "[$(Get-Date)] Validation file was created successfully."
        
    }
    'uninstall' {
        # Add Registry Values back
        New-ItemProperty -Path "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\" -Name syncwindowsettingdays -PropertyType DWORD -Value 3 -Force -Confirm:$false | Out-Null
        New-ItemProperty -Path "Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\cached mode\" -Name syncwindowsetting -PropertyType DWORD -Value 0 -Force -Confirm:$false | Out-Null
            
        # Delete validation file
        Remove-Item -Path "$validation\Outlook Remove Exchange Cache Silder Policy.txt" -Force -Confirm:$false -ErrorAction SilentlyContinue

        # Log the result of the uninstallation
        Add-Content -Path $logfile -Value "[$(Get-Date)] Outlook Remove Exchange Cache Silder Policy script version $($version) was uninstalled successfully"

        }
    }
