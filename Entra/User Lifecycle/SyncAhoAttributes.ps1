# Clean Up Data between Azure and Aho

# Authenticate with Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

# Read the CSV file
$csvPath = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\NameImport.csv"
$users = Import-Csv -Path $csvPath

# Create an array to store updated user information
$updatedUsers = @()

# Loop through each user in the CSV
foreach ($user in $users) {
    # Fetch user details from Microsoft Graph using displayName
    $graphUsers = Get-MgbetaUser -Filter "displayName eq '$($user.displayName)'"

    # Check if user is found and is a member (not a guest)
    foreach ($graphUser in $graphUsers) {
        if ($graphUser -and $graphUser.UserType -eq "Member") {
            # Add UserPrincipalName and ObjectId to the user object
            $user | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $graphUser.UserPrincipalName
            $user | Add-Member -MemberType NoteProperty -Name "ObjectId" -Value $graphUser.Id

            # Add the updated user to the array
            $updatedUsers += $user
            break # Break the loop after finding the first matching member user
        }
    }
    
    # Check if no matching member user was found
    if (-not ($user.PSObject.Properties.Name -contains "UserPrincipalName")) {
        Write-Host "Member user not found: $($user.displayName)"
        # Add the original user to the array without updates
        $updatedUsers += $user
    }
}

# Export the updated user list to a new CSV file
$newCsvPath = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\NameImport2.csv"
$updatedUsers | Export-Csv -Path $newCsvPath -NoTypeInformation

# Disconnect from Microsoft Graph
Disconnect-MgGraph


#Sync Attributes across to Azure from Aho


$ImportCSV = Import-Csv -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\19Feb2024.xlsx"

$ImportCSV | ForEach-Object {
    
    $Attributes = @{
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate' = $_.START_DATE
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime' = $_.PROJECTED_END_DATE
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'= $_.EMPLOYEE_TYPE
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $_.ASSIGNMENT_CATEGORY
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = $_.TEAM_GROUP
    }

    Update-MgBetaUser -UserId $_.ObjectID -AdditionalProperties $Attributes


    Write-Host "$($_.UserPrincipalName) has been updated with attributes `n $($attributes)"

}