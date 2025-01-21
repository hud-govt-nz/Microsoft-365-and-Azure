try {
    $DefaultPrinter = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default='TRUE'"
    
    if ($DefaultPrinter -eq $null) {
        Write-Output "No default printer found."
        Exit 1
    }
    
    if ($DefaultPrinter.Name -ine "HUD Print Anywhere") {
        Write-Output "Default Printer is not HUD Print Anywhere, updating default printer..."
        Exit 1
    }
    else {
        Write-Output "HUD Print Anywhere is set as default printer, no action required."
        Exit 0
    }
}
catch {
    $errMsg = $_.Exception.Message
    Write-Output $errMsg
    Exit 1
}