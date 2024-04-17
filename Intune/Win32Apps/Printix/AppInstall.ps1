<#
.APP: Printix Client
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
$AppName=[string]'Printix Client'
$AppVersion="1.3.1254.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"
$AppInstallFile= "CLIENT_{hud.printix.net}_{13ef4486-2503-4ca9-86c4-1d7b5fae76d7}.MSI"
$AppInstallArguments='WRAPPED_ARGUMENTS=/id:13ef4486-2503-4ca9-86c4-1d7b5fae76d7:oms /qn'

# Check if Application Already Exists
$Installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString
$Installed += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString

$Result = @()
foreach ($item in $AppName) {
    $tempResult = $Installed | Where-Object { $_.DisplayName -ne $null } | Where-Object { $_.DisplayName -match $item }
    $Result += @($tempResult)
    }

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        if ($Result[0]) {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName is currently installed with version $AppVersion."
            } else {
                try {
                    #  Copy Installer to local device
                    Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
                    $outfile = "$Path\$AppInstallFile"
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error copying installer: $_"
                        exit 1
                        }
                try {
                    # Install App
                    $installProcess = Start-Process $outfile -ArgumentList $AppInstallArguments -PassThru -Wait -ErrorAction Stop
                    $exitCode = $installProcess.ExitCode

                    # Log the result of the installation
                    if ($exitCode -eq 0) {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was installed successfully with exit code $($exitCode)"
                        } else {
                            Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was not installed successfully with exit code $($exitCode)" 
                            exit $exitCode
                            }
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
                    Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction Stop
                    Write-Log -Path $AppLog -Value "[$(Get-Date)] Installer files were deleted successfully."
                    } catch {
                        Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting installer files: $_"
                        exit 1
                        }
                }
        }

        'uninstall' {
            if ($Result) {
                try {
                    # Uninstall App
                    $uninstall_command = "C:\Program Files\printix.net\Printix Client\unins000.exe"
                    $uninstall_args = '/VERYSILENT'
                    $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
                    $exitCode = $uninstallProcess.ExitCode
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
    

    