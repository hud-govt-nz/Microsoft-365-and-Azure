$ENGB = Get-AppxPackage -allusers *LanguageExperiencePacken-GB*
try {
    if ($ENGB -ne $null){
    Write-Host "En-GB Language Pack found, removing"    
    Exit 1
    }
    else {
    Write-Host "en-GB is not installed on this device"
    Exit 0
    }
} catch {
    $errMsg = $_.exeption.essage
    Write-Output $errMsg
    }

