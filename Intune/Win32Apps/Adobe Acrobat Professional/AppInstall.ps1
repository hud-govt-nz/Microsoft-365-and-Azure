<#
.APP: Adobe Acrobat (64-bit)
.AUTHOR: Ashley Forde
.DATE: 16 May 2023
#>

# this is a test line to show change control 2
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
$AppName=[string]'Adobe Acrobat (64-bit)'
$AppVersion="23.001.20174"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"
$AppInstallFile= "$path\APRO23.0\Adobe Acrobat\setup.exe"
$AppInstallArguments='/S'

# Check if Application Already Exists
$Search = 'Adobe Acrobat*', 'Adobe Creative Cloud*'
$Installed = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString
$Installed += Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, UninstallString

$Result = @()
foreach ($item in $Search) {
    $tempResult = $Installed | Where-Object { $_.DisplayName -ne $null } | Where-Object { $_.DisplayName -match $item }
    $Result += @($tempResult)
    }

# Separate Adobe Acrobat and Adobe Creative Cloud
$AdobeAcrobat = $Result | Where-Object { $_.DisplayName -match 'Adobe Acrobat*' }
$AdobeCreativeCloud = $Result | Where-Object { $_.DisplayName -match 'Adobe Creative Cloud*' }

# Check if the install or uninstall switch is used
switch ($args[0]) {
    'install' {
        if ($Result[0]) {
            Write-Log "Acrobat Professional is currently installed with version $($AdobeAcrobat.DisplayVersion), Uninstalling..."
            try {
                # Uninstall Old App
                $uninstall_args = (($Result[0].UninstallString -split ' ')[1] -replace '/I','/X ') + ' /q'
                $uninstallProcess = Start-Process MsiExec.exe -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
                $exitCode = $uninstallProcess.ExitCode
                Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was uninstalled successfully with exit code $($exitCode)"
                } catch {
                    Write-Log "Error uninstalling Acrobat Professional: $_"
                    exit 1
                    }

        if ($AdobeCreativeCloud) {
            try {
                # Uninstall Adobe Creative Cloud Desktop Client
                $uninstall_command = $AdobeCreativeCloud.UninstallString
                $uninstall_args = ' -uninstall '
                $uninstallProcess = Start-Process $uninstall_command -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
                $exitCode = $uninstallProcess.ExitCode
                Write-Log "Acrobat Creative Cloud Desktop Client $($AdobeCreativeCloud.DisplayVersion) was uninstalled successfully with exit code $exitCode"
                } catch {
                    Write-Log "Error uninstalling Adobe Creative Cloud: $_"
                    exit 1
                    }
            }
            try {
                # Copy Installer to local device
                Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
                $outfile = $AppInstallFile
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
            } else {
                try {
                    #  Copy Installer to local device
                    Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
                    $outfile = $AppInstallFile
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
        if ($Result[0]) {
            try {
                # Uninstall App
                $uninstall_args = (($Result[0].UninstallString -split ' ')[1] -replace '/I','/X ') + ' /q'
                $uninstallProcess = Start-Process MsiExec.exe -ArgumentList $uninstall_args -PassThru -Wait -ErrorAction Stop
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
