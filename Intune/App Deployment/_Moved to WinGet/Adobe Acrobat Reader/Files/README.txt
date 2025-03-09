Download URL: https://get.adobe.com/uk/reader/enterprise/

1. Download the latest vesion from the URL above
2. extract the .exe using 7-zip.
3. add files to .\Files directory
4. Amend the existing setup.ini file to include the text below

Update the Setup.ini file

[Startup]
RequireMSI=3.0
CmdLine=/sall /rs

[Product]
PATCH=AcroRdrDCUpd2400220759.msp
msi=AcroRead.msi
CmdLine=TRANSFORMS="AcroRead.mst"
