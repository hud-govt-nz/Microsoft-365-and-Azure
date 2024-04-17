<#
.APP: HUD - Custom Shortcuts
.AUTHOR: Ashley Forde
.DATE: 16 May 2023
#>

# Root Folder
$Directory = 'Tools'

# Define Log function
function Write-Log {
    param(
        [string]$Path,
        [string]$Value
        )
    Add-Content -Path $Path -Value $Value
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
# Get Current User
$User= (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]

#Set application details
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'HUD - Custom Shortcuts'
$AppVersion="2.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"

# Define the file patterns to be removed
$filePatterns = @("*Aho.lnk*", "*HUD Support Hub*")

# Define the paths to be scanned
$paths = @(
    "C:\Users\Public\Desktop\",
    "C:\Users\$User\OneDrive - Ministry of Housing and Urban Development\Desktop",
    "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Log",
    "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Validation",
    "C:\Users\$User\AppData\Roaming\HUD Tools\EndpointManager\Data"
)

try {
    # Loop over each path
    foreach ($path in $paths) {
        if (Test-Path $path) {
            # Write the path to the log
            Write-Log -Path $AppLog -Value "[$(Get-Date)] Scanning path: $path"

            # Loop over each file pattern
            foreach ($filePattern in $filePatterns) {
                # Write the file pattern to the log
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Looking for pattern: $filePattern"

                # Get the child items of the path and filter by the file pattern
                Get-ChildItem -Path $path | Where-Object { $_.Name -ilike $filePattern } | ForEach-Object {
                    # Write the file removal to the log
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Removing file: $($_.FullName)"
                    Remove-Item $_ -Force -Recurse -Confirm:$false -ErrorAction Stop
                }
            }
        }
        else {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] Path not found: $path"
        }
    }
}
catch {
    Write-Host "An error occurred: $_"
    Write-Log -Path $AppLog -Value "[$(Get-Date)] An error occurred: $_"
}

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        try {
            #  Copy HUD Custom Shortcuts Icon files to local device
            $IcoFiles = New-Item -Path $HomeFolder -Name 03_Icons -ItemType Directory -Force -Confirm:$false
            $LinkFiles = New-Item -Path $HomeFolder -Name 04_Links -ItemType Directory -Force -Confirm:$false
            Copy-Item -Path ".\Installer\*" -Destination $IcoFiles -Recurse -Force
            
            # Create an array of shortcut names and URLs
            $shortcuts = @(
                @{name="Aho"; url="https://fa-evjy-saasfaprod1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseWelcome"; icon="$IcoFiles\Aho.ico"},
                @{name="HUD Support Hub"; url="https://mhud.sharepoint.com/sites/im"; icon="$IcoFiles\HUD.ico"}
                )
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error copying files: $_"
                exit 1
                }                
        try {
            #Create Shortcuts
            foreach($shortcut in $shortcuts) {
                $link = "$($LinkFiles)\" + $shortcut.name + ".lnk"
                [System.IO.Path]::GetFileName($link)
                
                # Create shortcut that can be pinned to taskbar
                $wshShell = New-Object -ComObject WScript.Shell
                $objShortcut = $wshShell.CreateShortcut($link)
                $objShortcut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                $objShortcut.Arguments = $shortcut.url
                $objShortcut.IconLocation = "$($shortcut.icon)"
                $objShortcut.Save()
    
                Copy-Item -Path $link -Destination "C:\Users\Public\Desktop" -Force -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log -Path $AppLog -Value "[$(Get-Date)] $($shortcut.name) shortcut successfully installed"
                }
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error creating shortcuts: $_"
                exit 1
                }
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
            try {
                # Loop over each path
                foreach ($uninstallPath in $paths) {
                    if (Test-Path $uninstallPath) {
                        # Write the path to the log
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Scanning path: $uninstallPath"
        
                        # Loop over each file pattern
                        foreach ($filePattern in $filePatterns) {
                            # Write the file pattern to the log
                            Write-Log -Path $AppLog -Value "[$(Get-Date)] Looking for pattern: $filePattern"
        
                            # Get the child items of the path and filter by the file pattern
                            Get-ChildItem -Path $uninstallPath | Where-Object { $_.Name -ilike $filePattern } | ForEach-Object {
                                # Write the file removal to the log
                                Write-Log -Path $AppLog -Value "[$(Get-Date)] Removing file: $($_.FullName)"
                                Remove-Item $_.FullName -Force -Recurse -Confirm:$false -ErrorAction Stop    
                            }   
                        }
                        Remove-Item -Path $HomeFolder\03_Icons -Force -Recurse -Confirm:$false -ErrorAction SilentlyContinue
                        Remove-Item -Path $HomeFolder\04_Links -Force -Recurse -Confirm:$false -ErrorAction SilentlyContinue  
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Successfully removed folders 03_Icons and 04_Links"
                    } else {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Path not found: $uninstallPath"    
                    }
                }
            } catch {
                Write-Host "An error occurred: $_"
                Write-Log -Path $AppLog -Value "[$(Get-Date)] An error occurred: $_"
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
        
    default {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Invalid argument. Please specify 'install' or 'uninstall'."
        exit 1
    }
}
