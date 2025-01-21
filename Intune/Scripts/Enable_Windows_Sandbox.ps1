<#
.SYNOPSIS
    Enable Windows Sandbox.
.DESCRIPTION
    This script will simply enable the Windows Sandbox for the user.
.NOTES
#>

#Set variables as input for the script
$WindowsFeature = "Containers-DisposableClientVM"

try {
	#Get the state of the Windows Feature
	$WindowsFeatureState = (Get-WindowsOptionalFeature -FeatureName $WindowsFeature -Online).State
}
catch {
	Write-Error "Failed to get the state of $WindowsFeature"
}

#Verify if the Windows Feature is enabled
if ($WindowsFeatureState -eq "Enabled") {
	Write-Output "$WindowsFeature is enabled"
}
else {
	try {
		#Enable the Windows Feature
		Enable-WindowsOptionalFeature -FeatureName $WindowsFeature -Online -NoRestart -ErrorAction Stop
		Write-Output "Successfully enabled $WindowsFeature"
	}
	catch {
		Write-Error "Failed to enable $WindowsFeature"
	}
}
