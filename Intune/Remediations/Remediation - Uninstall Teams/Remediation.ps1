if ($null -eq (Get-AppxPackage -Name MicrosoftTeams -AllUsers)) {
    Write-Output "Microsoft Teams Personal App not present"
    Exit 0
} else {
    try {
        Write-Output "Removing Microsoft Teams Personal App"
        
        if (Get-Process -Name "msteams" -ErrorAction SilentlyContinue) {
            try {
                Write-Output "Stopping Microsoft Teams Personal app process"
                Stop-Process -Name "msteams" -Force
                Write-Output "Stopped"
            } catch {
                Write-Output "Unable to stop process, trying to remove anyway"
            }
        }
        
        Get-AppxPackage -Name MicrosoftTeams -AllUsers | Remove-AppxPackage -AllUsers
        Write-Output "Microsoft Teams Personal App removed successfully"
        Exit 0
    } catch {
        Write-Error "Error removing Microsoft Teams Personal App"
        Exit 1
    }
}
