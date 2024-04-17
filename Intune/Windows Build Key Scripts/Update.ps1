# Install the NuGet package provider if not already installed
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Force -Confirm:$False | Out-Null

# Install the PSWindowsUpdate module if not already installed
Install-Module -Name PSWindowsUpdate -Scope AllUsers -Force -Confirm:$False | Out-Null

# Import the PSWindowsUpdate module
Import-Module -Name PSWindowsUpdate

Write-Host "Downloading and installing Windows updates."

# Get updates and install them
Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Verbose

# Shutdown 
Write-Host "Shutting down in 10 Seconds..."

Shutdown /s /t 10