Function Remove-Certificates {
    <#
    .SYNOPSIS
    This function removes certificates from the local certificate store.
    .PARAMETER Thumbprint
    The thumbprint of the certificate to be removed.
    .PARAMETER OutputPath
    The output path for exporting the certificate information.
    .NOTES
    This function is strongly typed to ensure proper data handling.
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$Thumbprint,

        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    <#
    .RETURN
    Boolean value indicating the success of the removal operation.
    #>
    try {
        # Search by Thumbprint
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$OLDCA = Get-ChildItem cert:\ -Recurse | where{$_.Thumbprint -eq $Thumbprint} 
        if ($OLDCA -eq $Null) {
            throw "Certificates with thumbprint $Thumbprint not found in the local certificate store."
        }

        # Export .CSV of cert info to the specified output path
        $OLDCA | Export-Csv -Path $OutputPath -NoTypeInformation

        # Remove all certificates with the specified thumbprint from the local certificate store
        $OLDCA | ForEach-Object {Remove-Item $_.PSPath}
        Write-Host "All certificates with thumbprint $Thumbprint removed successfully."
        return $true
    }
    catch {
        Write-Host $_
        return $false
    }
}

#Run through and remove certificates on device
Remove-Certificates -Thumbprint 4d0e6d3b4bb391c8625433046545378d21a548ce -OutputPath C:\HUD\export_certificates2.csv