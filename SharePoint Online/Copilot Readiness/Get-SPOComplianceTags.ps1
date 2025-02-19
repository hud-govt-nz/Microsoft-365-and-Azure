#=============================================================================
# Script Name: Get-SPOComplianceTags.ps1
# Created: 2024
# Author: GitHub Copilot
# 
# Description:
#   This script retrieves and displays all compliance tags (retention labels) that 
#   have been published to SharePoint Online sites. It provides information about
#   available labels that can be applied to content.
#
# Parameters:
#   -SiteUrl : String
#       Optional. Specify a specific site URL to check. If not provided,
#       checks the root site collection.
#
# Example Usage:
#   # Check root site collection
#   .\Get-SPOComplianceTags.ps1
#
#   # Check specific site
#   .\Get-SPOComplianceTags.ps1 -SiteUrl "https://mhud.sharepoint.com/sites/example"
#=============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$SiteUrl = "https://mhud.sharepoint.com"
)

# Disable PnP PowerShell update check
$env:PNPPOWERSHELL_UPDATECHECK = "Off"

# Initialize basic configuration
Write-Host "Initializing script..." -ForegroundColor Cyan

# Connect to SharePoint site
try {
    Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -ClientId $env:DigitalSupportAppID -Tenant 'mhud.onmicrosoft.com' -Thumbprint $env:DigitalSupportCertificateThumbprint

    # Get all compliance tags using PnP
    Write-Host "`nRetrieving compliance tags..." -ForegroundColor Yellow
    
    # Query for available compliance labels
    $complianceTags = Get-PnPLabel

    if ($complianceTags) {
        Write-Host "`nFound $($complianceTags.Count) published compliance tags:" -ForegroundColor Green
        
        $complianceTags | ForEach-Object {
            Write-Host "`nTag Details:" -ForegroundColor Cyan
            Write-Host "  Name: $($_.TagName)" -ForegroundColor White
            Write-Host "  ID: $($_.TagID)" -ForegroundColor White
            if ($_.Description) {
                Write-Host "  Description: $($_.Description)" -ForegroundColor White
            }
            Write-Host "  Is Enabled: $($_.Enabled)" -ForegroundColor White
            
            # Add a separator line
            Write-Host ("=" * 80) -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "`nNo compliance tags found for this site." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Disconnect the PnP connection
    Disconnect-PnPOnline
    Write-Host "`nScript completed." -ForegroundColor Green
}