<#
.SYNOPSIS
    Camtasia 2024.

.DESCRIPTION
    Script to install Camtasia and to purge previous installs
    
.NOTES
    - AUTHOR: Ashley Forde
    - Version: 2.0
    - Date: 12.02.2024
#>

#Region Parameters
[CmdletBinding()]
param(
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[ValidateSet("Install","Uninstall")]
	[string]$Mode
)
# EndRegion Parameters
# Region Functions
function Write-LogEntry {
	param(
		[Parameter(Mandatory = $true,HelpMessage = "Value added to the log file.")]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[Parameter(Mandatory = $true,HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("1","2","3")]
		[string]$Severity,
		[Parameter(Mandatory = $false,HelpMessage = "Name of the log file that the entry will written to.")]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = $LogFileName
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $logsFolderVar -ChildPath $FileName

	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff")," ",(Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))

	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")

	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"

	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
		if ($Severity -eq 1) {
			Write-Verbose -Message $Value
		}
		elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	}
	catch [System.Exception]{
		Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}

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
		# Force creating 00_Staging folder at a minimum if it is missing
		New-Item -Path "$HomeFolder" -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false | Out-Null

	}
	else {
		Write-Verbose "Creating root folder..."
		New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false
		if (-not $?) {
			Write-Verbose "Failed to create $HomeFolder"
		}

		# Create subfolders
		foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
			New-Item -Path "$HomeFolder\" -Name $subFolder -ItemType "directory" -Force -Confirm:$false
			if (-not $?) {
				Write-Verbose -Message "Failed to create sub-folder $subFolder under $HomeFolder"
			}
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
# EndRegion Functions

# Comment: This region contains initialisations and variable assignments required for the script.   
# Region Initialisations
$HomeFolder = "C:\HUD"
$folderPaths = Initialize-Directories -HomeFolder $HomeFolder
# EndRegion Initialisations

# Assign the returned values to individual variables
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$validationFolderVar = $folderPaths.ValidationFolder

# Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$AppName = "Camtasia_2024"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "2.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
	Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}

$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName
Write-LogEntry -Value "Setup folder has been created at: $Setupfolder." -Severity 1

# Install/Uninstall M365 Apps
if ($Mode -eq "Install") {

    #Clean up previous version install
    try {
        $MyApp = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Camtasia*" }
    
        if ($null -ne $MyApp) {
            $MyApp.Uninstall()
            Write-LogEntry -Value "Camtasia has been uninstalled successfully." -Severity 1

            # Remove Previous Validation File
            try {
                # Delete validation file
                Remove-Item -Path "C:\HUD\02_Validation\Camtasia.txt" -Force -ErrorAction Stop
                Write-LogEntry -Value "Previous Validation file was deleted successfully." -Severity 1
                } catch {
                    Write-LogEntry -Value "Error deleting validation file: $_" -Severity 3
                    exit 1
                    }
        } else {
            Write-LogEntry -Value "Camtasia is not installed on this system." -Severity 1
        }
    } catch {
        Write-LogEntry -Value "An error occurred during the uninstallation process: $_" -Severity 3
    }

	try {
		# Copy files to staging folder
		Copy-Item -Path "$PSScriptRoot\Installer\*" -Destination $SetupFolder -Recurse -Force -ErrorAction Stop
		Write-LogEntry -Value "Installer files have been copied to $Setupfolder." -Severity 1

		# Test if file(s) exists
		$PreReqSetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "VC_redist.x64.exe").ToString()
		if (-not (Test-Path $PreReqSetupFilePath)) { throw "Error: Setup file not found" }

		$SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "Camtasia.msi").ToString()
		if (-not (Test-Path $SetupFilePath)) { throw "Error: Setup file not found" }

		Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

		try {
			# Check setup.exe has valid file signature
			$PreReqVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($PreReqSetupFilePath).FileVersion
			Write-LogEntry -Value "Setup is running version $PreReqVersionInfo" -Severity 1

			try {
				# Run Pre Req Software MS Visual C++ 2019 Redistributable Package
				Write-LogEntry -Value "Starting $Mode of MS Visual C++ 2019 Redistributable Package" -Severity 1
				[string]$Arguments0 = "/q /norestart"
				$Process0 = Start-Process $PreReqSetupFilePath -ArgumentList $Arguments0 -Wait -PassThru -ErrorAction Stop
				
				# Post Install Actions
				if ($Process0.ExitCode -eq "0") {
					Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
				} else {
					Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process0.ExitCode)" -Severity 3
				}

				# Run  setup.exe per configuration file
				Write-LogEntry -Value "Starting $Mode of Camtasia 2024" -Severity 1
				[string]$Arguments = "TRANSFORMS=$SetupFolder\camtasia.mst /quiet /norestart"
				$Process = Start-Process $SetupFilePath -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

				# Post Install Actions
				if ($Process.ExitCode -eq "0") {
					# Create validation file
					New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
					Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
					Write-LogEntry -Value "Install of $AppName is complete" -Severity 1
				} else {
					Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3
				}

				# Cleanup 
				if (Test-Path "$SetupFolder") {
					Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
					Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
				}

			} catch {
				Write-LogEntry -Value "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
				return # Stop execution of the script after logging a critical error
			}
		} catch [System.Exception]{ Write-LogEntry -Value "Error preparing installation $FileName $($mode). Errormessage: $($_.Exception.Message)" -Severity 3 }

	} catch [System.Exception]{ Write-LogEntry -Value "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3 }

}
elseif ($Mode -eq "Uninstall") {

    try {
        $MyApp = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Camtasia*" }
    
        if ($null -ne $MyApp) {
            $MyApp.Uninstall()
            Write-LogEntry -Value "Camtasia has been uninstalled successfully." -Severity 1

            # Delete validation file
            try {
                Remove-Item -Path "C:\HUD\02_Validation\Camtasia_2024.txt" -Force -ErrorAction Stop
                Write-LogEntry -Value "Validation file was deleted successfully." -Severity 1
                } catch {
                    Write-LogEntry -Value "Error deleting validation file: $_" -Severity 3
                    exit 1
                    }

			Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1

        } else {
            Write-LogEntry -Value "Camtasia is not installed on this system." -Severity 1
        }

		# Cleanup 
		if (Test-Path "$SetupFolder") {
			Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
			Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
		}
		
    } catch [System.Exception]{ 
        Write-LogEntry -Value "Error completing uninstall. Errormessage: $($_.Exception.Message)" -Severity 3
    }
} 
#ENDS
