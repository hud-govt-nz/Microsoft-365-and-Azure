$subjectName = Read-Host -Prompt "Please enter certificate subject name"
$certStore = "LocalMachine"
$validityPeriod = 24
$certFolder = Read-Host -Prompt "Please enter the location to save the certificate"

# Define the certificate parameters
$certParams = @{
    Subject           = "CN=$subjectName"
    CertStoreLocation = "Cert:\$certStore\My"
    KeyExportPolicy   = "Exportable"
    KeySpec           = "Signature"
    NotAfter          = (Get-Date).AddMonths($validityPeriod)
}

# Create the self-signed certificate
$cert = PKI\New-SelfSignedCertificate @certParams

# Define the export parameters for the public and private keys

$certExportParams = @{
    Cert     = $cert
    FilePath = "$certFolder\$subjectName.cer"
}
$pfxExportParams = @{
    Cert         = "Cert:\$certStore\My\$($cert.Thumbprint)"
    FilePath     = "$certFolder\$subjectName.pfx"
    ChainOption  = "EndEntityCertOnly"
    NoProperties = $true
    Password     = (Read-Host -Prompt "Enter password for your certificate" -AsSecureString)
}

# Export the public and private keys
Export-Certificate @certExportParams
Export-PfxCertificate @pfxExportParams
