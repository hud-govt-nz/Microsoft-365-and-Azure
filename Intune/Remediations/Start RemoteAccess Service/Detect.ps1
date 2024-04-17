# Detect
$RemoteAccess = Get-service -Name RemoteAccess
 
if ($RemoteAccess.Status -eq "Stopped") {  
    Exit 1 # Commence remediation

    } else {
        Exit 0 # Service is already Active

        }