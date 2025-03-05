<#
.SYNOPSIS
    Deploys and configures Always On VPN connection for HUD.

.DESCRIPTION
    This script creates and configures an Always On VPN connection with the following features:
    - Removes any existing VPN connections
    - Configures EAP authentication
    - Sets up split tunneling
    - Configures automatic connection
    - Adds necessary routes
    - Enables DNS registration

.NOTES
    Version: 1.0
    Author: Ashley Forde
    Last Modified: 4 March 2025
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Install','Uninstall')]
    [string]$Mode
)

# Reference functions.ps1
. "$PSScriptRoot\functions.ps1"

# Application Variables
$AppName = "HUD ALO VPN - User Context"
$AppVersion = "1.0"

# Initialize Directories
$folderpaths = Initialize-Directories -HomeFolder C:\HUD\

# Template Variables
$Date = Get-Date -Format "MM-dd-yyyy"
$stagingFolderVar = $folderPaths.StagingFolder
$logsFolderVar = $folderPaths.LogsFolder
$LogFileName = "$($AppName)_${Mode}_$Date.log"
$validationFolderVar = $folderPaths.ValidationFolder
$AppValidationFile = "$validationFolderVar\$AppName.txt"

# Begin Setup
Write-LogEntry -Value "Initiating setup process" -Severity 1

# Define constants
$vpnName = "HUD ALO VPN"
$serverAddress = "azuregateway-f6dd002a-af3b-467c-a8ac-669542f8d9ea-00b2af831edc.vpn.azure.com"

function Remove-ExistingVpnConnection {
    Write-LogEntry -Value "Checking for existing VPN connections..." -Severity "1"
    $existingVPNs = @()
    
    try {
        $existingVPNs += Get-VpnConnection -AllUserConnection -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq $vpnName }
        $existingVPNs += Get-VpnConnection -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq $vpnName }
    } catch {
        Write-LogEntry -Value "No VPN connections found or unable to query VPN connections" -Severity "1"
        return
    }

    if ($existingVPNs.Count -gt 0) {
        Write-LogEntry -Value "Found $($existingVPNs.Count) existing VPN connection(s) named '$vpnName'" -Severity "1"
        foreach ($vpn in $existingVPNs) {
            try {
                if ($vpn.AllUserConnection) {
                    Write-LogEntry -Value "Removing All-User VPN connection: $($vpn.Name)" -Severity "1"
                    Remove-VpnConnection -Name $vpn.Name -AllUserConnection -Force -ErrorAction Stop
                } else {
                    Write-LogEntry -Value "Removing User-Level VPN connection: $($vpn.Name)" -Severity "1"
                    Remove-VpnConnection -Name $vpn.Name -Force -ErrorAction Stop
                }
            } catch {
                if ($_.Exception.Message -like "*was not found*") {
                    Write-LogEntry -Value "VPN connection already removed or not found: $($vpn.Name)" -Severity "1"
                } else {
                    Write-LogEntry -Value "Failed to remove VPN connection $($vpn.Name): $_" -Severity "2"
                }
                continue
            }
        }

        # Verify removals were successful
        Start-Sleep -Seconds 2
        $remainingVPNs = @()
        $remainingVPNs += Get-VpnConnection -AllUserConnection -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq $vpnName }
        $remainingVPNs += Get-VpnConnection -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq $vpnName }
        
        if ($remainingVPNs.Count -gt 0) {
            Write-LogEntry -Value "Some VPN connections may still exist" -Severity "2"
        } else {
            Write-LogEntry -Value "Successfully removed all existing VPN connections named '$vpnName'" -Severity "1"
        }
    } else {
        Write-LogEntry -Value "No existing VPN connections found with name '$vpnName'" -Severity "1"
    }
}

function Remove-NetworkProfiles {
    Write-LogEntry -Value "Cleaning up network profiles from registry..." -Severity "1"
    $profilePaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles",
        "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Signatures\Unmanaged"
    )

    foreach ($path in $profilePaths) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path | ForEach-Object {
                $profileDescription = Get-ItemProperty -Path $_.PSPath -Name "Description" -ErrorAction SilentlyContinue
                if ($profileDescription.Description -like "*$vpnName*") {
                    Write-LogEntry -Value "Found matching network profile: '$($profileDescription.Description)'" -Severity "1"
                    try {
                        Remove-Item -Path $_.PSPath -Force -ErrorAction Stop
                        Write-LogEntry -Value "Successfully removed network profile '$($profileDescription.Description)' from $path" -Severity "1"
                    } catch {
                        Write-LogEntry -Value "Failed to remove network profile '$($profileDescription.Description)' from $path : $_" -Severity "2"
                    }
                }
            }
        } else {
            Write-LogEntry -Value "Registry path not found: $path" -Severity "1"
        }
    }
}

function Set-VpnAutoConnect {
    param([bool]$Enable)
    
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Rasman\Parameters\Configs"
    if ($Enable) {
        if (!(Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        $vpnStrategy = @{
            "VpnStrategy" = 7
            "AlwaysOn" = 1
            "DeviceTunnel" = 0
        }
        foreach ($key in $vpnStrategy.Keys) {
            New-ItemProperty -Path $registryPath -Name $key -Value $vpnStrategy[$key] -PropertyType DWORD -Force
        }
    } else {
        if (Test-Path $registryPath) {
            Remove-Item -Path $registryPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-VpnDnsRegistration {
    param([bool]$Enable)
    
    $path = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Parameters"
    if ($Enable) {
        Set-ItemProperty -Path $path -Name "RegisterDNS" -Value 1 -Type DWord
    } else {
        Remove-ItemProperty -Path $path -Name "RegisterDNS" -ErrorAction SilentlyContinue
    }
}

function Remove-VpnRoutes {
    $routes = @(
        "10.0.0.0/22",
        "10.0.4.0/22",
        "10.1.0.0/16",
        "20.36.104.6/32",
        "20.36.104.7/32",
        "20.36.105.32/29",
        "20.53.48.96/27",
        "20.36.112.6/32",
        "20.36.113.0/32",
        "20.36.113.32/29",
        "29.53.48.96/27"
    )

    foreach ($route in $routes) {
        try {
            Remove-VpnConnectionRoute -ConnectionName $vpnName -DestinationPrefix $route -ErrorAction Stop
            Write-LogEntry -Value "Removed route: $route" -Severity "1"
        } catch {
            Write-LogEntry -Value "Failed to remove route $route : $_" -Severity "2"
        }
    }
}

function Set-PbkAlwaysOnCapable {
    param(
        [string]$VpnName
    )
    
    Write-LogEntry -Value "Setting AlwaysOnCapable in PBK file..." -Severity "1"
    $pbkPath = "$env:ProgramData\Microsoft\Network\Connections\Pbk\rasphone.pbk"
    
    if (Test-Path $pbkPath) {
        $content = Get-Content $pbkPath
        $vpnSection = $false
        $alwaysOnFound = $false
        $newContent = @()
        
        foreach ($line in $content) {
            $newContent += $line
            
            if ($line -match "^\[.*$VpnName\]$") {
                $vpnSection = $true
            }
            elseif ($line -match "^\[.*\]$") {
                if ($vpnSection -and -not $alwaysOnFound) {
                    $newContent += "AlwaysOnCapable=1"
                }
                $vpnSection = $false
                $alwaysOnFound = $false
            }
            elseif ($vpnSection -and $line -match "^AlwaysOnCapable=") {
                $newContent[-1] = "AlwaysOnCapable=1"
                $alwaysOnFound = $true
            }
        }
        
        if ($vpnSection -and -not $alwaysOnFound) {
            $newContent += "AlwaysOnCapable=1"
        }
        
        $newContent | Set-Content $pbkPath
        Write-LogEntry -Value "PBK file updated successfully" -Severity "1"
    }
    else {
        Write-LogEntry -Value "PBK file not found at $pbkPath" -Severity "2"
    }
}

function Backup-HostsFile {
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    $backupPath = "$hostsPath.bak"
    
    Write-LogEntry -Value "Creating backup of hosts file..." -Severity "1"
    try {
        Copy-Item -Path $hostsPath -Destination $backupPath -Force
        Write-LogEntry -Value "Successfully created hosts file backup at $backupPath" -Severity "1"
    } catch {
        Write-LogEntry -Value "Failed to create hosts file backup: $_" -Severity "3"
        throw
    }
}

function Update-HostsFileEntries {
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    
    # Define new entries
    $newEntries = @(
        "# Azure Private Endpoints"
        "10.0.4.10    property.database.windows.net"
        "10.0.4.30    sql-fpdreporting-dev.database.windows.net"
        "10.0.4.40    sql-reporting-prod.database.windows.net"
        "10.0.5.5     dlprojectsdataprod.blob.core.windows.net"
        "10.0.5.6     dlreportingdataprod.blob.core.windows.net"
        "10.0.5.10    dlreportingdataprod.dfs.core.windows.net"
    )

    try {
        # Read existing content
        $content = Get-Content -Path $hostsPath -ErrorAction Stop
        
        # Filter out existing entries
        $cleanedContent = $content | Where-Object { 
            $_ -notmatch "Azure Private Endpoints" -and
            $_ -notmatch "property\.database\.windows\.net" -and
            $_ -notmatch "sql-fpdreporting-dev\.database\.windows\.net" -and
            $_ -notmatch "sql-reporting-prod\.database\.windows\.net" -and
            $_ -notmatch "dlprojectsdataprod\.blob\.core\.windows\.net" -and
            $_ -notmatch "dlreportingdataprod\.blob\.core\.windows\.net" -and
            $_ -notmatch "dlreportingdataprod\.dfs\.core\.windows\.net"
        }

        # Combine cleaned content with new entries
        $updatedContent = @()
        $updatedContent += $cleanedContent
        $updatedContent += ""
        $updatedContent += $newEntries

        # Create a temporary file
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Write content to temp file first
        $updatedContent | Out-File -FilePath $tempFile -Encoding ASCII -Force
        
        # Use Copy-Item with Force to replace the hosts file
        Copy-Item -Path $tempFile -Destination $hostsPath -Force
        
        # Clean up temp file
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        
        Write-LogEntry -Value "Successfully updated hosts file entries" -Severity "1"
    } catch {
        Write-LogEntry -Value "Failed to update hosts file: $_" -Severity "3"
        throw
    }
}

# Add/Remove
switch ($Mode) {
    "Install" {
        Write-LogEntry -Value "Installing HUD ALO VPN..." -Severity "1"
        
        Backup-HostsFile
        Update-HostsFileEntries
        
        Remove-ExistingVpnConnection
        Remove-NetworkProfiles
        Set-VpnAutoConnect -Enable $true
        
        Write-LogEntry -Value "Creating new VPN connection..." -Severity "1"
        $EAP = '<EapHostConfig
xmlns="http://www.microsoft.com/provisioning/EapHostConfig">
    <EapMethod>
        <Type
            xmlns="http://www.microsoft.com/provisioning/EapCommon">13
        </Type>
        <VendorId
            xmlns="http://www.microsoft.com/provisioning/EapCommon">0
        </VendorId>
        <VendorType
            xmlns="http://www.microsoft.com/provisioning/EapCommon">0
        </VendorType>
        <AuthorId
            xmlns="http://www.microsoft.com/provisioning/EapCommon">0
        </AuthorId>
    </EapMethod>
    <Config
        xmlns="http://www.microsoft.com/provisioning/EapHostConfig">
        <Eap
            xmlns="http://www.microsoft.com/provisioning/BaseEapConnectionPropertiesV1">
            <Type>13</Type>
            <EapType
                xmlns="http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV1">
                <CredentialsSource>
                    <CertificateStore>
                        <SimpleCertSelection>true</SimpleCertSelection>
                    </CertificateStore>
                </CredentialsSource>
                <ServerValidation>
                    <DisableUserPromptForServerValidation>true</DisableUserPromptForServerValidation>
                    <ServerNames></ServerNames>
                    <TrustedRootCA>9F EC AE 14 B7 2F C7 16 F7 36 56 EE CD 14 14 DE A0 86 2C DD </TrustedRootCA>
                </ServerValidation>
                <DifferentUsername>false</DifferentUsername>
                <PerformServerValidation
                    xmlns="http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV2">false
                </PerformServerValidation>
                <AcceptServerName
                    xmlns="http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV2">true
                </AcceptServerName>
                <TLSExtensions
                    xmlns="http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV2">
                    <FilteringInfo
                        xmlns="http://www.microsoft.com/provisioning/EapTlsConnectionPropertiesV3">
                        <CAHashList Enabled="true">
                            <IssuerHash>D1 48 48 EF 2E B5 BE F9 BA 7B 70 D0 8A 06 3A 53 B2 FA B3 AA </IssuerHash>

                        </CAHashList>
                    </FilteringInfo>
                </TLSExtensions>
            </EapType>
        </Eap>
    </Config>
</EapHostConfig>'

        $vpnConnection = Add-VpnConnection -Name $vpnName `
            -ServerAddress $serverAddress `
            -TunnelType Ikev2 -AuthenticationMethod Eap -SplitTunneling:$True `
            -RememberCredential -EncryptionLevel Required -PassThru `
            -EapConfigXmlStream $EAP -AllUserConnection

        # Set AlwaysOnCapable in PBK file
        Set-PbkAlwaysOnCapable -VpnName $vpnName

        # Configure VPN connection settings
        Write-LogEntry -Value "Configuring VPN connection settings..." -Severity "1"
        Set-VpnConnection -Name $vpnName `
            -SplitTunneling $True `
            -RememberCredential $True -AllUserConnection

        # Enable Auto Connect in RasMan config
        Write-LogEntry -Value "Enabling auto-connect..." -Severity "1"
        $ConnectionType = "rasentry"
        $ConnectionName = $vpnName
        $Path = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Config\$ConnectionType\$ConnectionName"

        if(!(Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }

        New-ItemProperty -Path $Path -Name "AutoTrigger" -Value 2 -PropertyType DWORD -Force

        # Output VPN connection details
        Write-LogEntry -Value "VPN Connection Details:" -Severity "1"
        Write-LogEntry -Value "Name: $($vpnConnection.Name)" -Severity "1"
        Write-LogEntry -Value "GUID: $($vpnConnection.GUID)" -Severity "1"
        Write-LogEntry -Value "Server: $($vpnConnection.ServerAddress)" -Severity "1"
        Write-LogEntry -Value "Connection Status: $($vpnConnection.ConnectionStatus)" -Severity "1"

        Set-VpnDnsRegistration -Enable $true
        
        $routes = @(
            "10.0.0.0/22",
            "10.0.4.0/22",
            "10.1.0.0/16",
            "20.36.104.6/32",
            "20.36.104.7/32",
            "20.36.105.32/29",
            "20.53.48.96/27",
            "20.36.112.6/32",
            "20.36.113.0/32",
            "20.36.113.32/29",
            "29.53.48.96/27"
        )

        foreach ($route in $routes) {
            try {
                Add-VpnConnectionRoute -ConnectionName $vpnName -DestinationPrefix $route -ErrorAction Stop
                Write-LogEntry -Value "Added route: $route" -Severity "1"
            } catch {
                Write-LogEntry -Value "Failed to add route $route : $_" -Severity "2"
            }
        }
        
        Write-LogEntry -Value "VPN installation completed successfully." -Severity "1"
        
        # Create validation file
        try {
            New-Item -ItemType File -Path $AppValidationFile -Force -Value $AppVersion
            Write-LogEntry -Value "Validation file has been created at $AppValidationFile" -Severity "1"
        } catch {
            Write-LogEntry -Value "Failed to create validation file: $_" -Severity "3"
        }
    }
    "Uninstall" {
        Write-LogEntry -Value "Uninstalling HUD ALO VPN..." -Severity "1"
        
        # Restore original hosts file if backup exists
        $hostsPath = "$env:windir\System32\drivers\etc\hosts"
        $backupPath = "$hostsPath.bak"
        if (Test-Path $backupPath) {
            Copy-Item -Path $backupPath -Destination $hostsPath -Force
            Remove-Item -Path $backupPath -Force
            Write-LogEntry -Value "Restored original hosts file" -Severity "1"
        }
        
        Remove-ExistingVpnConnection
        Remove-NetworkProfiles
        Remove-VpnRoutes
        Set-VpnAutoConnect -Enable $false
        Set-VpnDnsRegistration -Enable $false
        
        # Clean up RasMan config
        $ConnectionType = "rasentry"
        $Path = "HKLM:\SYSTEM\CurrentControlSet\Services\RasMan\Config\$ConnectionType\$vpnName"
        if (Test-Path $Path) {
            Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue
        }
        
        # Remove validation file if it exists
        if (Test-Path $AppValidationFile) {
            try {
                Remove-Item -Path $AppValidationFile -Force
                Write-LogEntry -Value "Validation file removed successfully" -Severity "1"
            } catch {
                Write-LogEntry -Value "Failed to remove validation file: $_" -Severity "2"
            }
        }
        
        Write-LogEntry -Value "VPN uninstallation completed successfully." -Severity "1"
    }
    default {
        Write-LogEntry -Value "Invalid mode: $Mode" -Severity "3"
    }
}