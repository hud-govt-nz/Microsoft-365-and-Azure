
<#
.APP: Te Reo Maori Addin (Outlook)
.DESCRIPTION: For use within Outlook, provides a helpful dropdown menu with various formal, and informal greeintgs and closing lines for email.
.AUTHOR: Ashley Forde c/f MBIE Digital
.DATE: 13 Sept 23
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
$AppName=[string]'Te Reo Maori Addin'
$AppVersion="1.0"
$AppValidationFile="$validation\$AppName.txt"
$AppLog="$logs\$AppName.log"
$AppInstallFile= "Setup.msi"
$AppInstallArguments="/qn"

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
            # Copy Installer to local device
            Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
            $outfile = "$Path\$AppInstallFile"

            # Copy Greetings CSV to ProgramData
            New-Item -Name TeReoAddin -Path "C:\ProgramData\" -ItemType Directory -Force -Confirm:$false
            Copy-Item -Path "$Path\TeReoGreetings.csv" -Destination "C:\ProgramData\TeReoAddin\"
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
        # Copy Installer to local device
        Copy-Item -Path ".\Installer\*" -Destination $path -Recurse -Force
        $outfile = "$Path\$AppInstallFile"

        # Unnstall App
        $installProcess = Start-Process "Msiexec.exe" -ArgumentList @("/X", "$outfile", "/qn") -PassThru -Wait -ErrorAction Stop
        $exitCode = $installProcess.ExitCode

        # Clear Te Reo Maori Addin Registry Key
        $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'
        $searchFor = 'Te Reo Maori Addin'

        # Search the registry
        $foundKeys = Get-ChildItem -Path $registryPath | ForEach-Object {
            $key = $_
            $properties = Get-ItemProperty -Path $key.PSPath

            # Check if the DisplayName matches the search criteria
            if ($properties.DisplayName -eq $searchFor) {
                $key.PSPath
            }
        }

        # Display or remove found keys
        if ($null -ne $foundKeys) {
            $foundKeys | ForEach-Object {
                Write-Log -Path $AppLog -Value "[$(Get-Date)] Found matching key at $_"

                # Uncomment the following line to remove the key; caution, this is irreversible!
                Remove-Item -Path $_ -Recurse
            }
        } else {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] No matching keys found."
          }

        # Log the result of the installation
        if ($exitCode -eq 0) {
          Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was uninstalled successfully with exit code $($exitCode)"
          } else {
              Write-Log -Path $AppLog -Value "[$(Get-Date)] $AppName version $AppVersion was not uninstalled with exit code $($exitCode)" 
              exit $exitCode
              }
        } catch {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] Error installing App: $_"
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
      try {
        # Delete installer files
        Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction Stop
        Remove-Item -Path "C:\ProgramData\TeReoAddin\" -Recurse -Force -ErrorAction Stop
        Write-Log -Path $AppLog -Value "[$(Get-Date)] Installer files were deleted successfully."
        } catch {
            Write-Log -Path $AppLog -Value "[$(Get-Date)] Error deleting installer files: $_"
            exit 1
            }
      }
    }

  default {
    Write-Log -Path $AppLog -Value "[$(Get-Date)] Invalid argument. Please specify 'install' or 'uninstall'."
    exit 1
    }
  }
