Connect-MgGraph

$ImportExcel = Import-Excel -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Book3.xlsx"

$ImportExcel | ForEach-Object {

    $UserPrincipalName = $_.Email
    $Group = $_.Group

    $attributes = @{
	'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = $Group
        }

	Update-MgBetaUser -UserId $UserPrincipalName -AdditionalProperties $attributes
	Write-Host "User: $UserPrincipalName`nOrganisational Group: $Group" -ForegroundColor Green
}