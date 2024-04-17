Connect-MicrosoftTeams

$Users = Import-CSV -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Documents\AD-AAD Change 24102022\Sync DDI Phone No to Users Again\DDI List from AD.csv"

$Users | ForEach-Object {

    Write-Host "Updating $($_.UserPrincipalName)"

    Try {
        Set-CsPhoneNumberAssignment -Identity $_.UserPrincipalName -PhoneNumber $_.OfficePhone -PhoneNumberType DirectRouting -ErrorAction Stop
        Grant-CsTenantDialPlan -Identity $_.UserPrincipalName -PolicyName DP-04Region  -ErrorAction Stop
        Grant-CsOnlineVoiceRoutingPolicy -Identity $_.UserPrincipalName -PolicyName VP-Unrestricted  -ErrorAction Stop

        Write-Host "Number and Policies applied to $($_.UserPrincipalName)" -ForegroundColor Green
    
    }Catch{
        
        Write-Host 'Updates Failed for this user' -ForegroundColor Red
    }

     
}
