# Import the CSV file
$Path = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Name&MGRName.xlsx"
$Names = Import-Excel -Path $Path

# Initialize an array to store the updated data
$UpdatedData = @()

# Iterate over each row in the CSV
foreach ($NameRow in $Names) {
    # Retrieve the UserPrincipalName for NAME
    $NameUPN = (Get-MgUser -Filter "displayName eq '$($NameRow.NAME)'").UserPrincipalName

    # Retrieve the UserPrincipalName for MGRName
    $MgrNameUPN = (Get-MgUser -Filter "displayName eq '$($NameRow.MGRName)'").UserPrincipalName

    # Add the retrieved UserPrincipalNames to the current row
    $NameRow | Add-Member -MemberType NoteProperty -Name "NAMEUPN" -Value $NameUPN
    $NameRow | Add-Member -MemberType NoteProperty -Name "MGRNAMEUPN" -Value $MgrNameUPN

    # Add the updated row to the array
    $UpdatedData += $NameRow
}

# Export the updated data to the CSV file
$UpdatedData | Export-Excel $Path -AutoSize -AutoFilter -WorksheetName "Names" -FreezeTopRow -BoldTopRow


