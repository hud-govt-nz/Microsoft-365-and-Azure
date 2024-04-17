<#
.APP: HUD - New Deveice Setup
.AUTHOR: Ashley Forde
.DATE: 16 May 2023
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
$AppName=[string]'HUD - New Device Setup'
$AppVersion="2.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"

try {
    # App Package Removal
    $Appx =@{
        0 = "XboxIdentityProvider"
        1 = "XboxSpeechToTextOverlay"
        2 = "XboxGamingOverlay"
        3 = "XboxGameOverlay"
        4 = "Xbox.TCUI"
        5 = "MicrosoftTeams"
        }
    
    Foreach ($Key in $Appx.Keys) {
        try {
            Get-AppxPackage -Name "*$($Appx[$Key])*" -AllUsers | Remove-AppPackage -AllUsers -Confirm:$false -ErrorAction Stop -Verbose
            Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) has been successfully uninstalled."
            } catch {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) has not been successfully uninstalled."
                }
        }

    # Enabling .NET 3.5
    if((Get-WindowsOptionalFeature -FeatureName NetFx3 -Online).State -ne "Enabled") {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) NetFx3 is currently Disabled, enabling Component."
        Enable-WindowsOptionalFeature -FeatureName NetFx3 -Online
        } else {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) NetFx3 is currently Enabled, Skipping Installation."
            }

    # Add Serial to Settings Menu
    $Path0 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"
    if((Get-ItemProperty -Path $Path0).SerialNumberIsValid -ne 1) {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) SerialNumberIsValid is not present. Adding Key & Value to show Serial Number on Settings page."
        Set-ItemProperty -Path $Path0 -Name SerialNumberIsValid -Value 1 -Force
        } else {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] $($Appx[$Key]) SerialNumberIsValid is already present on device. No action will be taken."
            
        }
    } catch {
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error copying installer: $_"
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