$TenantID = "9e9b3020-3d38-48a6-9064-373bc7b156dc"
$Input = "C:\HUD\06_Reporting\steam_token.txt"
$output = "C:\HUD\06_Reporting\Stream"

.\StreamClassicVideoReportGenerator_V1.14.ps1 -AadTenantId $TenantID -InputFile $Input -OutDir $output -Verbose