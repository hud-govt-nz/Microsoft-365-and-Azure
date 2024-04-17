$AzureAD = Get-InstalledModule -Name AzureAD -MinimumVersion 2.0.2.140 -ErrorAction SilentlyContinue

#Detect if the AzureAD module is installed, action according to state
try {
    if ($Null -eq $AzureAD.Name) {
    
    Install-PackageProvider -Name NuGet -Scope AllUsers -Force -Confirm:$false 
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Install-Module -Name AzureAD -Scope AllUsers -Repository PSGallery -Force -Confirm:$false 
    Write-Output "Azure AD Module has been installed on this device."
    exit 0
    }

    else
    {
    Write-Output "Azure AD module is detected, no action required."
    exit 0
    }
} 
 catch{
    $errMsg = $_.exeption.essage
    Write-Output $errMsg
 }
 #endregion