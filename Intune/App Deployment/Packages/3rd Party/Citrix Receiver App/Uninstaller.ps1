function Initialize-Directories {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$HomeFolder
	)

	# Check if the path exists
	if (Test-Path -Path $HomeFolder) {
		Write-Verbose "Home folder exists..."

		# Check if subfolders exist, if not create them
		foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
			$subFolderPath = Join-Path -Path $HomeFolder -ChildPath $subFolder
			if (-not (Test-Path -Path $subFolderPath)) {
				Write-Verbose "Creating sub-folder $subFolder..."
				New-Item -Path $HomeFolder -Name $subFolder -ItemType "directory" -Force -Confirm:$false
			}
			elseif ($subFolder -eq "00_Staging") {
				Write-Verbose "Emptying 00_Staging folder..."
				Remove-Item -Path $subFolderPath\* -Recurse -Force -Confirm:$false
			}
		}

	}
	else {
		Write-Verbose "Creating root folder..."
		New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false

		# Create subfolders
		foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
			Write-Verbose "Creating sub-folder $subFolder..."
			New-Item -Path $HomeFolder -Name $subFolder -ItemType "directory" -Force -Confirm:$false
		}
	}

	# Calculate subfolder paths
	$StagingFolder = Join-Path -Path $HomeFolder -ChildPath "00_Staging"
	$LogsFolder = Join-Path -Path $HomeFolder -ChildPath "01_Logs"
	$ValidationFolder = Join-Path -Path $HomeFolder -ChildPath "02_Validation"

	# Return the folder paths as a custom object
	return @{
		HomeFolder = $HomeFolder
		StagingFolder = $StagingFolder
		LogsFolder = $LogsFolder
		ValidationFolder = $ValidationFolder
	}
}

function Get-InstalledApps {
    param (
        [string[]]$App
    )

    $Installed = Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' }
    $Installed += Get-ItemProperty -Path HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' }
	$Installed += Get-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' }
	$Installed += Get-ItemProperty -Path HKCU:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.DisplayName -ne '' }

    $SelectedApp = @()
    foreach ($item in $App) {
        $tempResult = $Installed | Where-Object { $_.DisplayName -match $item }
        $SelectedApp += @($tempResult)
    }

    return $SelectedApp | Select-Object -First 1
}

function Write-LogEntry {
    param (
        [string]$Value,
        [int]$Severity
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [Severity $Severity] $Value"
    Add-Content -Path (Join-Path -Path $script:logsFolderVar -ChildPath $script:LogFileName) -Value $logEntry
}

function Stop-CitrixProcesses {
    Write-LogEntry -Value "Starting Citrix process termination" -Severity 1
    $citrixProcesses = @(
        "Receiver",
        "concentr",
        "wfcrun32",
        "redirector",
        "AuthManSvr",
        "picaMain",
        "viewer",
        "CDViewer",
        "concentr",
        "wfica32",
        "CitrixWorkspaceApp",
        "CitrixReceiverUpdater",
        "SelfServicePlugin",
        "CitrixCCMEngine"
    )

    foreach ($proc in $citrixProcesses) {
        $processes = Get-Process -Name $proc -ErrorAction SilentlyContinue
        if ($processes) {
            Write-LogEntry -Value "Stopping $proc processes..." -Severity 1
            $processes | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Additional cleanup - Remove Citrix services
    $citrixServices = Get-Service -Name "Citrix*" -ErrorAction SilentlyContinue
    foreach ($service in $citrixServices) {
        Write-LogEntry -Value "Stopping and removing service: $($service.Name)" -Severity 1
        Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
        $service | Set-Service -StartupType Disabled -ErrorAction SilentlyContinue
    }
    
    Start-Sleep -Seconds 2  # Give processes time to fully terminate
    Write-LogEntry -Value "Completed Citrix process termination" -Severity 1
}

function Uninstall-Application {
    param (
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )

    try {
        $maxAttempts = 3
        $attempt = 1
        $success = $false

        while (-not $success -and $attempt -le $maxAttempts) {
            Write-LogEntry -Value "Attempt $attempt of $maxAttempts to uninstall $AppName" -Severity 1
            
            $app = Get-InstalledApps -App $AppName
            if ($app) {
                # Configure silent switches based on common installers
                $silentSwitches = @{
                    'msiexec.exe' = '/qn /norestart'
                    'setup.exe'   = '/silent /quiet'
                    'TrolleyExpress.exe' = '/silent /uninstall /cleanup /force /noreboot'
                    'default'     = '/S /s /Q /q /quiet /silent /SILENT /VERYSILENT /noreboot'
                }

                # Try different uninstall methods in order of preference
                $uninstallMethods = @(
                    @{ Type = "QuietUninstall"; String = $app.QuietUninstallString },
                    @{ Type = "MSIUninstall"; String = "msiexec.exe /x $($app.PSChildName) $($silentSwitches['msiexec.exe'])" },
                    @{ Type = "StandardUninstall"; String = $app.UninstallString }
                )

                foreach ($method in $uninstallMethods) {
                    if ($method.String) {
                        Write-LogEntry -Value "Trying $($method.Type) method for $AppName" -Severity 1
                        
                        # Modify the uninstall string to ensure silent operation
                        $uninstallCmd = $method.String
                        
                        # Handle MSI uninstalls
                        if ($uninstallCmd -match "msiexec") {
                            $uninstallCmd = $uninstallCmd -replace "/I", "/X" -replace "/i", "/x"
                            if ($uninstallCmd -notmatch "/qn") {
                                $uninstallCmd = "$uninstallCmd $($silentSwitches['msiexec.exe'])"
                            }
                        }
                        # Handle Citrix TrolleyExpress
                        elseif ($uninstallCmd -match "TrolleyExpress\.exe") {
                            if ($uninstallCmd -notmatch "/silent") {
                                $uninstallCmd = "$uninstallCmd $($silentSwitches['TrolleyExpress.exe'])"
                            }
                        }
                        # Handle other installers
                        else {
                            $hasQuietSwitch = $false
                            foreach ($switch in $silentSwitches['default'].Split(' ')) {
                                if ($uninstallCmd -match [regex]::Escape($switch)) {
                                    $hasQuietSwitch = $true
                                    break
                                }
                            }
                            if (-not $hasQuietSwitch) {
                                $uninstallCmd = "$uninstallCmd $($silentSwitches['default'])"
                            }
                        }

                        Write-LogEntry -Value "Executing uninstall command: $uninstallCmd" -Severity 1
                        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait -PassThru -WindowStyle Hidden -ErrorAction SilentlyContinue

                        if ($process.ExitCode -eq 0) {
                            Write-LogEntry -Value "Successfully uninstalled $AppName using $($method.Type)" -Severity 1
                            $success = $true
                            break
                        } elseif ($process.ExitCode -eq 3010 -or $process.ExitCode -eq 1641) {
                            Write-LogEntry -Value "Uninstallation of $AppName requires reboot (Exit code: $($process.ExitCode))" -Severity 2
                            $success = $true
                            $script:IntuneExitCode = 3010
                            break
                        } else {
                            Write-LogEntry -Value "Failed to uninstall $AppName. Exit code: $($process.ExitCode)" -Severity 3
                        }
                    }
                }

                if (-not $success) {
                    Write-LogEntry -Value "Attempting registry cleanup for $AppName" -Severity 2
                    # Additional cleanup if all methods fail
                    Write-Host "Attempting registry cleanup..."
                    $regPaths = @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
                    )

                    foreach ($path in $regPaths) {
                        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DisplayName -match $AppName } | 
                        ForEach-Object { 
                            Remove-Item $_.PSPath -Force -Recurse -ErrorAction SilentlyContinue
                        }
                    }
                }
            } else {
                Write-LogEntry -Value "Application $AppName not found in registry" -Severity 1
                break
            }

            if (-not $success) {
                $attempt++
                if ($attempt -le $maxAttempts) {
                    Write-LogEntry -Value "Retrying uninstallation after cleanup..." -Severity 2
                    Stop-CitrixProcesses
                    Start-Sleep -Seconds 5
                }
            }
        }
    } catch {
        Write-LogEntry -Value ("Error uninstalling {0}: {1}" -f $AppName, $_.Exception.Message) -Severity 3
        $script:IntuneExitCode = 1
    }
}

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Script Variables
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "Citrix_Uninstall.log"
$logsFolderVar = $folderPaths.LogsFolder
$IntuneExitCode = 0

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $logsFolderVar)) {
    New-Item -Path $logsFolderVar -ItemType Directory -Force | Out-Null
}

Write-LogEntry -Value "Starting Citrix uninstallation process" -Severity 1

# Get all Citrix applications installed on the machine
$citrixApps = Get-InstalledApps -App "Citrix"

# First stop all Citrix processes
Write-LogEntry -Value "Initiating Citrix process cleanup" -Severity 1
Stop-CitrixProcesses

# Uninstall each Citrix application found
if ($citrixApps) {
    foreach ($app in $citrixApps) {
        Write-LogEntry -Value "Found Citrix application: $($app.DisplayName)" -Severity 1
        Uninstall-Application -AppName $app.DisplayName
    }
} else {
    Write-LogEntry -Value "No Citrix applications found" -Severity 1
}

# Validate uninstallation
$remainingApps = Get-InstalledApps -App "Citrix"
if ($remainingApps) {
    Write-LogEntry -Value "Validation Failed: Citrix applications still present after uninstallation" -Severity 3
    $script:IntuneExitCode = 1
    
    # Create validation file with failed status
    $validationContent = @{
        Status = "Failed"
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        RemainingApps = $remainingApps.DisplayName
        ExitCode = $IntuneExitCode
    } | ConvertTo-Json

    $validationPath = Join-Path -Path $folderPaths.ValidationFolder -ChildPath "Citrix_Uninstall_Status.json"
    $validationContent | Out-File -FilePath $validationPath -Force
} else {
    Write-LogEntry -Value "Validation Successful: No Citrix applications remain installed" -Severity 1
    
    # Create validation file with success status
    $validationContent = @{
        Status = "Success"
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        ExitCode = $IntuneExitCode
    } | ConvertTo-Json

    $validationPath = Join-Path -Path $folderPaths.ValidationFolder -ChildPath "Citrix_Uninstall_Status.json"
    $validationContent | Out-File -FilePath $validationPath -Force
}

Write-LogEntry -Value "Citrix uninstallation process completed with exit code: $IntuneExitCode" -Severity 1

# Return exit code for Intune
exit $IntuneExitCode