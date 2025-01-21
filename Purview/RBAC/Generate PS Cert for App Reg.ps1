# Your tenant name (can something more descriptive as well)
$TenantName        = "mhud.onmicrosoft.com"

# Where to export the certificate without the private key
$CerOutputPath     = "C:\HUD\HUD Digital Support - Compliance PS Certificate.cer"

# What cert store you want it to be in
$StoreLocation     = "Cert:\CurrentUser\My"

# Expiration date of the new certificate
$ExpirationDate    = (Get-Date).AddYears(2)


# Splat for readability
$CreateCertificateSplat = @{
    Subject = 'CN=HUD Digital Support Compliance,O=MHUD, OU=MHUD,C=MHUD'
    FriendlyName      = "HUD Support CompliancePS"
    DnsName           = $TenantName
    CertStoreLocation = $StoreLocation
    NotAfter          = $ExpirationDate
    KeyExportPolicy   = "Exportable"
    KeySpec           = "KeyExchange"

}

$mycert = New-SelfSignedCertificate @CreateCertificateSplat


# Export certificate to .cer file
$mycert | Export-Certificate -FilePath HUD-Support-CompliancePS.cer