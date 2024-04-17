<#
.APP: HUD - Font Install
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

#Set application details
$path = "$HomeFolder\00_Staging"
$logs = "$HomeFolder\01_Logs"
$validation = "$HomeFolder\02_Validation"
$AppName=[string]'HUD - Font Install'
$AppVersion="2.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        try {
            # Copy Installer to local device
            Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force -Confirm:$false
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error copying files: $_"
                exit 1
                }                
        try {
            # Add font file names to an array
            $fonts =@()
            $fonts += Get-ChildItem -Path $path

            # Load each file and assign registry string value
            Foreach ($font in $fonts) {
                copy-Item -Path "$path\$($font)" -Destination "$env:windir\Fonts" -Force -Confirm:$false
                New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $font.Name -PropertyType String -Value $font.Name -Force -Confirm:$false
                }

            # Log the result of the installation
            Write-Log -Path $AppLog -Value "[$(Get-Date)] HUD National Fonts were installed successfully."
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error installing App: $_"
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
        try {
            # Delete installer files
            Remove-Item -Path "$path\*" -Recurse -Force -Confirm:$false -ErrorAction Stop 
            Write-Log -Path $AppLog -Value "[$(Get-Date)] font files were deleted successfully."
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting font files: $_"
                exit 1
                }
        }

    'uninstall' {
        # Complile National font list array
        $installedfonts=@()
        $installedfonts += Get-ChildItem -Path C:\Windows\Fonts | Where-Object {$_.Name -ilike "*National*"}

        if ($null -ne $installedfonts) {
            try {
                #Uninstall HUD Font Install silently
                Foreach ($font in $installedfonts) {
                    Remove-Item -Path "C:\Windows\Fonts\$font" -Force -Confirm:$false 
                    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" -Name $font.name -Force -Confirm:$false
                    # Log the result of the uninstall
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] HUD Font $($font.name) version $($version) was uninstalled successfully."
                }
                } catch {
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Error uninstalling fonts: $_"
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
