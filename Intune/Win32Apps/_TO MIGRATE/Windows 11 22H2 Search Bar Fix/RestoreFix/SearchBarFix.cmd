MD C:\HUD\00_Staging\22H2SearchFix
Copy "%~dp0*.reg" C:\HUD\00_Staging\22H2SearchFix /Y
PUSHD C:\HUD\00_Staging\22H2SearchFix
regedit.exe /s ActiveXfix.reg
@echo 1.0>C:\HUD\00_Staging\22H2SearchFix\Ver1.0.txt
Del C:\HUD\00_Staging\22H2SearchFix\*.reg