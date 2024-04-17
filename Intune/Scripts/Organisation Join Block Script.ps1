<#
.SYNOPSIS
  This script will prevent the "Allow my organization to manage my device" prompt when signing in.

.DESCRIPTION

.PARAMETER <Parameter_Name>

.INPUTS
  None

.OUTPUTS
  Transcript file will be located in $Env:Programfiles\My-Apps\EndpointManager\Log\$PackageName-install.log

.NOTES
  Version:        1.0
  Author:         Ashley Forde
  Creation Date:  08 August 2022
  Purpose/Change: Initial script development
  
.EXAMPLE
  Script is run in Intune.
#>

#Transcript:
$PackageName = "OrgJoinBlock"

$HUDIntuneInstall = "$Env:Programfiles\My-Apps\EndpointManager"
Start-Transcript -Path "$HUDIntuneInstall\Log\$PackageName-install.log" -Force -Append -Confirm:$False

##YOUR CODE HERE

#Setting registry key to block AAD Registration to 3rd party tenants. 
    $RegistryLocation = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin\"
    $keyname = "BlockAADWorkplaceJoin"

#Test if path exists and create if missing
    if (!(Test-Path -Path $RegistryLocation)){
    Write-Output "Registry location missing. Creating"
    New-Item $RegistryLocation | Out-Null
    }

#Force create key with value 1 
    New-ItemProperty -Path $RegistryLocation -Name $keyname -PropertyType DWord -Value 1 -Force | Out-Null
    Write-Output "Registry key set"

Stop-transcript

##END OF YOUR CODE
