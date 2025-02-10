$Folder = "$($env:homedrive)\HUD"
$validationFolder = Join-Path -Path $Folder -ChildPath "02_Validation"
$validationFile = Join-Path -Path $validationFolder -ChildPath "Citrix_Uninstall_Status.json"

if (Test-Path $validationFile) {
    try {
        $validationStatus = Get-Content -Path $validationFile | ConvertFrom-Json
        
        Write-Host "Citrix Uninstallation Validation Results:"
        Write-Host "======================================"
        Write-Host "Status: $($validationStatus.Status)"
        Write-Host "Timestamp: $($validationStatus.Timestamp)"
        Write-Host "Exit Code: $($validationStatus.ExitCode)"
        
        if ($validationStatus.Status -eq "Failed") {
            Write-Host "`nRemaining Citrix Applications:"
            $validationStatus.RemainingApps | ForEach-Object {
                Write-Host "- $_"
            }
            exit 1
        } else {
            Write-Host "`nAll Citrix applications have been successfully removed."
            exit 0
        }
    }
    catch {
        Write-Host "Error reading validation file: $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Host "Validation file not found. Uninstallation status unknown."
    exit 1
}
