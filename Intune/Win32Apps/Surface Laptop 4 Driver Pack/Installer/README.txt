##READ ME##

Download Installer file from the following link:
https://www.microsoft.com/en-us/download/details.aspx?id=102924

The lines in the AppInstall and Validiation script need to be updated to reflect the file name and version of the installer package as it changes over time. 

- Update the following lines in the AppInstall.ps1 file.

    $AppVersion="22000_23.042.26034.0"
    $AppInstallFile= "SurfaceLaptop4_Win11_22000_23.042.26034.0.msi"

- update the following line in the Validation.ps1 file

    $version = "22000_23.042.26034.0"
