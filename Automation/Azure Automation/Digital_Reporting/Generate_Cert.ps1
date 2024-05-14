# Your tenant name (can something more descriptive as well)
$TenantName        = "mhud.onmicrosoft.com"

# Where to export the certificate without the private key
$CerOutputPath     = "C:\HUD\DigitalReporter.cer"

# What cert store you want it to be in
$StoreLocation     = "Cert:\CurrentUser\My"

# Expiration date of the new certificate
$ExpirationDate    = (Get-Date).AddYears(2)


# Splat for readability
$CreateCertificateSplat = @{
    FriendlyName      = "AzureApp"
    DnsName           = $TenantName
    CertStoreLocation = $StoreLocation
    NotAfter          = $ExpirationDate
    KeyExportPolicy   = "Exportable"
    KeySpec           = "Signature"
    Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    HashAlgorithm     = "SHA256"
}

# Create certificate
$Certificate = PKI\New-SelfSignedCertificate @CreateCertificateSplat

# Get certificate path
$CertificatePath = Join-Path -Path $StoreLocation -ChildPath $Certificate.Thumbprint


# Export certificate without private key
PKI\Export-Certificate -Cert $CertificatePath -FilePath $CerOutputPath | Out-Null
$mypwd = ConvertTo-SecureString -String 'Petticoat4330' -Force -AsPlainText
PKI\Export-PfxCertificate -Cert $CertificatePath -FilePath "C:\HUD\PowerShellGraphCert.pfx" -Password $mypwd | Out-Null