#Generate Certificate = Dev-BuildToRent & Prod-BuildToRent

#$certname = "Dev-BuildToRent"
$certname = "Prod-BuildToRent"

$Params = @{
    Subject = "CN=$certname"
    CertStoreLocation = "Cert:\CurrentUser\My"
    KeyExportPolicy = 'Exportable'
    KeySpec = 'Signature'
    KeyLength = '2048'
    KeyAlgorithm = 'RSA'
    HashAlgorithm = 'SHA256'
    NotAfter = "$((Get-Date).AddYears(2))"
}

$cert = PKI\New-SelfSignedCertificate @Params

#Create Certificate Password
$mypwd = ConvertTo-SecureString -String <#Replace {myPassword}#> -Force -AsPlainText  ## Replace {myPassword}

#Export PFX Certificate
PKI\Export-PfxCertificate -Cert $cert -FilePath "C:\HUD\14_Certificates\$certname.pfx" -Password $mypwd   ## Specify your preferred location
