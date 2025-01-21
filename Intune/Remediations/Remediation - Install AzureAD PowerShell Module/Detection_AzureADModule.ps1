$AzureAD = Get-InstalledModule -Name AzureAD -MinimumVersion 2.0.2.140 -ErrorAction SilentlyContinue

#Detect if the AzureAD module is installed, throw exit action depending on state.
try {
    if ($Null -eq $AzureAD.Name){
    Write-Output "Azure AD Module is not detected on this device, installing..."
    Exit 1
    }

    else {
    Write-Output "Azure AD module is detected, no action required."
    Exit 0
}
}

 catch{
    $errMsg = $_.exeption.essage
    Write-Output $errMsg
 }