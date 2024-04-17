try {
    $DefaultPrinter = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default='TRUE'"
    
    if ($DefaultPrinter -eq $null) {
        Write-Output "No default printer found."
        (New-Object -ComObject WScript.Network).SetDefaultPrinter('HUD Print Anywhere')
        Write-Output "Default printer has been set to HUD Print Anywhere because no default was found."
        Exit 0
    }

    if ($DefaultPrinter.Name -ne "HUD Print Anywhere") {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter('HUD Print Anywhere')
        Write-Output "Default printer has been updated to HUD Print Anywhere."
        Exit 0
    } else {
        Write-Output "HUD Print Anywhere is set as default printer, no action required."
        Exit 0
    }
} catch {
    $errMsg = $_.Exception.Message
    Write-Output $errMsg
    Exit 1
}