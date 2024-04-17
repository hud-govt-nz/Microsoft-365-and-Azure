<#
Connect-AzureAD 
Connect-MicrosoftTeams
#>

Start-Transcript -Path "C:\Support\TeamsDDI.txt" -NoClobber -Append


#Obtain UPN
$User = Import-Csv -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\AD-AAD Change 24102022\Sync DDI Phone No to Users Again\DDI List from AD.csv"
$UPN = $user.UserPrincipalName


#Obtain DDI
Foreach ($name in $UPN) {
    #Obtain DDI from Azure AD
    $RawDDI = (Get-AzureADUser -ObjectId $name).TelephoneNumber
    
    Write-host "The DDI for user $($name) is $($RawDDI)"
    
    #Format DDI for teams import
    $DDIFormatted = $RawDDI -replace " ",""


    Write-Host "Importing into MS Teams..."

    #Update value in Teams Admin Console
    Set-CsPhoneNumberAssignment -Identity $name -PhoneNumber $DDIFormatted -PhoneNumberType DirectRouting

}


Stop-Transcript








