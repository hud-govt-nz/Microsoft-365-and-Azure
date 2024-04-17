<#
.APP: Disable Outlook Auto Complete
.AUTHOR: Ashley Forde
.DATE: 29 May 2023
#>

# Root Folder
$Directory = 'HUD'

# Define Log function
function Write-Log {
    param(
        [string]$Path,
        [string]$Value
        )
    Add-Content -Path $Path -Value $Value
    }

function Get-CurrentUserSID {
    # Set current user
    $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
    $keys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse

    foreach ($key in $keys) {
        if (($key.GetValueNames() | ForEach-Object { $key.GetValue($_) }) -match $CurrentUser) {
            $sid = $key.PSChildName
            break
        }
    }

    # SID for current user
    return $sid
}

# Create Directories
$HomeFolder = "$($env:homedrive)\$Directory"
    if (Test-Path -Path $HomeFolder) { 
        "Path exists!"
        } else { 
            "Creating root folder..."
            New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
            
            foreach($subFolder in "00_Staging", "01_Logs", "02_Validation") {
                New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false | Out-Null
                }
            }

#Set application details
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'HUD - Disable Outlook Auto Complete'
$AppVersion="2.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"
$SID = Get-CurrentUserSID

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        try {
            # Get Current Username
            $User= (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]

            #Auto Complete Registry Key
            $ShowAutoSug = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\preferences\
            $value = $ShowAutoSug.ShowAutoSug
            $path = $ShowAutoSug.pspath

            # Log the result of the installation
            if ($null -eq $value -or $value -eq "1") {
                #Block Auto Suggestions from appearing in outlook during email address lookup
                New-ItemProperty -Path $path -Name ShowAutoSug -PropertyType DWORD -Value 0 -Force -Confirm:$false
                
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Auto Complete has been updated successfully to 'disabled' on client"
                } else {
                    #Creating PreferencePath on device
                    New-Item -Path $path -Force -Confirm:$false
                    #Block Auto Suggestions from appearing in outlook during email address lookup
                    New-ItemProperty -Path $path -Name ShowAutoSug -PropertyType DWORD -Value 0 -Force -Confirm:$false
                    
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Auto Complete key does not exist. Creating with value '0'" 
                    }
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error setting registry values."
                exit 1
                }

            #Restart Outlook
            Get-Process Outlook* | Stop-Process -Force -Confirm:$false -ErrorAction SilentlyContinue

            #Clear Auto Complete Cache    
            Get-childitem -Path "C:\Users\$User\AppData\Local\Microsoft\Outlook\RoamCache" -Filter "*autocomplete*" | Remove-item -Force -Confirm:$False -ErrorAction SilentlyContinue

        try {
            # Create validation file
            New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion | Out-Null
            Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was created successfully."
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error creating validation file: $_"
                exit 1
                }
                
        }

    'uninstall' {
        if ($Result) {
            try {
                # Registry Value
                $ShowAutoSug = Get-ItemProperty Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\preferences\       
                Remove-ItemProperty -Path Registry::HKEY_USERS\$SID\Software\Policies\Microsoft\office\16.0\outlook\preferences\ -Name ShowAutoSug -Force -Confirm:$false

                Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was uninstalled successfully with exit code $($exitCode)"
                        
                } catch {
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Error uninstalling App: $_"
                    exit 1
                    }
                try {
                    # Delete validation file
                    Remove-Item -Path $AppValidationFile -Force -ErrorAction Stop
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Validation file was deleted successfully."
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting validation file: $_"
                        exit 1
                        }
            }
        }
    default {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Invalid argument. Please specify 'install' or 'uninstall'."
        exit 1
        }
}
