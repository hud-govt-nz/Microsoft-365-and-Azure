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
function Get-CurrentUserSID {
    $currentUser = (Get-Process -IncludeUserName -Name explorer | Select-Object -First 1 | Select-Object -ExpandProperty UserName).Split("\")[1]
    $profileListKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" -Recurse

    foreach ($profileListKey in $profileListKeys) {
        if (($profileListKey.GetValueNames() | ForEach-Object { $profileListKey.GetValue($_) }) -match $currentUser) {
            $sid = $profileListKey.PSChildName
            break
        }
    }

    return $sid
}
function Set-DisableRoamingSignatures {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true)]
		[ValidateSet('Add','Remove')]
		[string]$Action,

		[Parameter(Mandatory = $false)]
		[ValidateSet(0,1)]
		[int]$ValueData
	)

	$currentUserSID = Get-CurrentUserSID

	$Hive = "HKEY_USERS"
	$KeyPath = "SOFTWARE\Microsoft\Office\16.0\Outlook\Setup"
	$ValueName = "DisableRoamingSignatures"
	$ValueType = "DWORD"

	$registryPath = "$Hive\$currentUserSID\$KeyPath"

	if ($Action -eq "Add") {
		New-ItemProperty -Path "Registry::$registryPath" -Name $ValueName -PropertyType $ValueType -Value $ValueData -Force
	}
	elseif ($Action -eq "Remove") {
		Remove-ItemProperty -Path "Registry::$registryPath" -Name $ValueName -Force
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

    return $SelectedApp | select -First 1
}
# Function to get .NET version
function Get-DotNetVersion {
    $ReleaseKey = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -EA 0
    return $ReleaseKey.Release
}
# Download function
function Start-DownloadFile {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    begin {
        $WebClient = New-Object -TypeName System.Net.WebClient
    }
    process {
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        $WebClient.DownloadFile($URL, (Join-Path -Path $Path -ChildPath $Name))
    }
    end {
        $WebClient.Dispose()
    }
}
function Invoke-FileCertVerification {
	param(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$FilePath
	)
	# Get a X590Certificate2 certificate object for a file
	$Cert = (Get-AuthenticodeSignature -FilePath $FilePath).SignerCertificate
	$CertStatus = (Get-AuthenticodeSignature -FilePath $FilePath).Status
	if ($Cert) {
		#Verify signed by Microsoft and Validity
		if ($cert.Subject -match "O=Microsoft Corporation" -and $CertStatus -eq "Valid") {
			#Verify Chain and check if Root is Microsoft
			$chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
			$chain.Build($cert) | Out-Null
			$RootCert = $chain.ChainElements | ForEach-Object { $_.Certificate } | Where-Object { $PSItem.Subject -match "CN=Microsoft Root" }
			if (-not [string ]::IsNullOrEmpty($RootCert)) {
				#Verify root certificate exists in local Root Store
				$TrustedRoot = Get-ChildItem -Path "Cert:\LocalMachine\Root" -Recurse | Where-Object { $PSItem.Thumbprint -eq $RootCert.Thumbprint }
				if (-not [string]::IsNullOrEmpty($TrustedRoot)) {
					Write-LogEntry -Value "Verified setupfile signed by : $($Cert.Issuer)" -Severity 1
					return $True
				}
				else {
					Write-LogEntry -Value "No trust found to root cert - aborting" -Severity 2
					return $False
				}
			}
			else {
				Write-LogEntry -Value "Certificate chain not verified to Microsoft - aborting" -Severity 2
				return $False
			}
		}
		else {
			Write-LogEntry -Value "Certificate not valid or not signed by Microsoft - aborting" -Severity 2
			return $False
		}
	}
	else {
		Write-LogEntry -Value "Setup file not signed - aborting" -Severity 2
		return $False
	}
}

# test