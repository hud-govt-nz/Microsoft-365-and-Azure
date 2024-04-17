# Define the file paths and passwords
$originalPfxPath = "C:\Users\Ashley.Forde\Downloads\Printer Certificates\Level 6 South\Original\certificate-hud-7wq-l6-02-south.pfx"
$newPfxPath = "C:\Users\Ashley.Forde\Downloads\Printer Certificates\Level 6 South\Modified\certificate-hud-7wq-l6-02-south.pfx"
$originalPassword = "DlwB0KC5pafmNjmUDYH7bMxs"
$newPassword = "c8nRwTgrnPLd"

# Import the PFX file with the original password
$pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$pfx.Import($originalPfxPath, $originalPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)

# Export the PFX file with the new password
$exportFlags = [System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12
$bytes = $pfx.Export($exportFlags, $newPassword)
[System.IO.File]::WriteAllBytes($newPfxPath, $bytes)