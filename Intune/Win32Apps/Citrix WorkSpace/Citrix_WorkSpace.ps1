<#
.SYNOPSIS
    Citrix WorkSpace Deployment Script for HUD Environment.

.DESCRIPTION
    Script to install custom Citrix WorkSpace deployments using .XML files and parameter switches depending on scenario.

.PARAMETER Mode
    Sets the mode of operation. Supported modes are Install or Uninstall.

.EXAMPLE 
    powershell.exe -executionpolicy bypass -file .\Citrix_WorkSpace.ps1 -Mode Install
    powershell.exe -executionpolicy bypass -file .\Citrix_WorkSpace.ps1 -Mode Uninstall

.NOTES
    - AUTHOR: Ashley Forde
    - Version: 1.0
    - Quick Start guide: https://docs.citrix.com/en-us/tech-zone/build/tech-papers/citrix-workspace-app.html
    - Named install file to CitrixWorkspaceAppWeb.exe: https://docs.citrix.com/en-us/tech-zone/build/tech-papers/citrix-workspace-app.html 
    - Install parameters: https://docs.citrix.com/en-us/citrix-workspace-app-for-windows/install#install-parameters
    - Install download site: https://www.citrix.com/downloads/workspace-app/windows/workspace-app-for-windows-latest.html 

#>

#Region Parameters
[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Install", "Uninstall")]
    [string]$Mode
)
# EndRegion Parameters

# Error handling
$global:ErrorActionPreference = "Stop"
if($verbose){ $global:VerbosePreference = "Continue" }

# Region Functions
function Write-LogEntry {
    param (
        [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = $LogFileName
    )
    # Determine log file location
    $LogFilePath = Join-Path -Path $logsFolderVar -ChildPath $FileName
    	
    # Construct time stamp for log entry
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
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
    catch [System.Exception] {
        Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}
function Start-DownloadFile {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    Begin {
        # Construct WebClient object
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    Process {
        # Create path if it doesn't exist
        if (-not(Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }

        # Start download of file
        $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
    }
    End {
        # Dispose of the WebClient object
        $WebClient.Dispose()
    }
}

function Initialize-Directories {
    [CmdletBinding()]
    param (
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
        foreach ($subFolder in "00_Staging", "01_Logs", "02_Validation") {
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
        HomeFolder       = $HomeFolder
        StagingFolder    = $StagingFolder
        LogsFolder       = $LogsFolder
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
$AppName = "Citrix_WorkSpace_${Mode}"
$AppValidationFile = "$validationFolderVar\$AppName.txt"
$AppVersion = "1.0"
$LogFileName = "$($AppName)_${Mode}_$Date.log"

# Comment: The script initiates the install process, logs the initiation, and performs cleanup of the setup folder if it exists.
# Initate Install
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Attempt Cleanup of SetupFolder
if (Test-Path "$stagingFolderVar") {
    Remove-Item -Path "$($stagingFolderVar)\*" -Recurse -Force -ErrorAction SilentlyContinue
}

# Create setup folder
$SetupFolder = (New-Item -ItemType "directory" -Path $stagingFolderVar -Name $AppName -Force).FullName

# Copy installer to setup folder
Copy-Item -Path $PSScriptRoot'\Installer\*' -Destination $SetupFolder -Force -Recurse -Confirm:$false

# Install/Uninstall Citrix WorkSpace
if ($Mode -eq "Install") {

    try {
        # Test if file exists
        $SetupFilePath = (Join-Path -Path $SetupFolder -ChildPath "CitrixWorkspaceAppWeb.exe").ToString()
        if (-Not (Test-Path $SetupFilePath)) {Throw "Error: Setup file not found"}

        Write-LogEntry -Value "Setup file ready at $($SetupFilePath)" -Severity 1

        try {
            # Run setup per configuration file
            Write-LogEntry -Value "Starting Citrix Workspace app $Mode" -Severity 1
            [string]$Arguments = "/silent /includeSSON /EnableCEIP=false /AutoUpdateStream=LTSR /FORCE_LAA=1"

            $Process = Start-Process $SetupFilePath -ArgumentList $Arguments -PassThru -ErrorAction Stop
            Wait-Process -InputObject $process           

            if ($mode -eq "Install") {
                # Add Validation File
                if ($Process.ExitCode -eq "0") {
                    New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
                    Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity 1
                    Write-LogEntry -Value "Citrix Workspace app was installed successfully $($Process.ExitCode)" -Severity 1
                    } else { Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3}
            } elseif ($Mode -eq "Uninstall")  {
                if ($Process.ExitCode -eq "0") {
                    Remove-Item -Path $AppValidationFile -Force -Confirm:$false
                    Write-LogEntry -Value "Validation file has been removed at $AppValidationFile" -Severity 1
                    Write-LogEntry -Value "Uninstall of $AppName is complete" -Severity 1
                    } else { Write-LogEntry -Value "Install of $AppName failed with ExitCode: $($Process.ExitCode)" -Severity 3}
                }
            # Cleanup 
            if (Test-Path "$SetupFolder") {
                Remove-Item -Path "$SetupFolder" -Recurse -Force -ErrorAction Continue
                Write-LogEntry -Value "Cleanup completed successfully" -Severity 1
                }

            } catch [System.Exception] {
                Write-LogEntry -Value  "Error running installer. Errormessage: $($_.Exception.Message)" -Severity 3
                return  # Stop execution of the script after logging a critical error
                }

    } catch [System.Exception] {
        Write-LogEntry -Value  "Error finding setup.exe Possible download error. Errormessage: $($_.Exception.Message)" -Severity 3
        }

}
elseif ($Mode -eq "Uninstall") {
 ########
}