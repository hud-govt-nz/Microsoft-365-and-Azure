$Shortcuts = Get-ChildItem "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\EasiPay.url" -ErrorAction SilentlyContinue

if ($Shortcuts) {
    Write-host "EasiPay.url found. Trigger Remediation script to uninstall"
    Exit 1
    } else {
        Write-host "EasiPay.url not found. Computer is compliant"
        Exit 0 
        }