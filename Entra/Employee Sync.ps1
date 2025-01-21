#______________________________________________________________________________________________________________________________________________________
# Requires -Modules Microsoft.Graph.Authentication
#______________________________________________________________________________________________________________________________________________________

# Function to connect to Microsoft Graph
try {
    Connect-MgGraph `
        -ClientId $env:DigitalSupportAppID `
        -TenantId $env:DigitalSupportTenantID `
        -CertificateThumbprint $env:DigitalSupportCertificateThumbprint `
        -NoWelcome
    Write-Host "Connected to Graph" -ForegroundColor Green

    $CollectToken = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users" -ContentType "txt" -OutputType HttpResponseMessage
    $Token        = $CollectToken.RequestMessage.Headers.Authorization.Parameter
    $Token | Out-Null
    } catch {
        Write-Host "Error connecting to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
        exit 1
}

#______________________________________________________________________________________________________________________________________________________

# Import the CSV file
#$Path = "C:\HUD\06_Reporting\Test_User_Me.xlsx"
$Path = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Phase 3 Changes\Sync_Data.xlsx"
$Names = Import-Excel -Path $Path

#______________________________________________________________________________________________________________________________________________________


# Loop through each user in Spreadsheet
$Names | ForEach-Object {
    
    # Variables
    $AZURE_ID = $_."Azure Object ID"
    $JOB_TITLE = $_."Position Name"
    $DEPARTMENT = $_."Department"
    $ORG_GROUP = $_.Group
    $START_DATE = $_."Start Date"
    $POSITION_TYPE = $_."Position Type"
    $ASSIGNMENT_CATEGORY = $_."Assignment Category"
    $LOCATION = $_.LOCATION
    $MANAGER_OBJECT_ID = $_."Manager Azure Object ID"
    
    Write-Host "Processing user with Azure ID: $AZURE_ID" -ForegroundColor Yellow

    # Assign Manager
    $MANAGER = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$MANAGER_OBJECT_ID"}
    Write-Output "Assigning manager with ID: $MANAGER_ID to user: $AZURE_ID"
    Set-MgUserManagerByRef -UserId $AZURE_ID -BodyParameter $MANAGER

    # Set Employee ID and OrgData
    $EmployeeOrgData = @{
        'division' = $ORG_GROUP
    }

    $attributes = @{
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType' = $POSITION_TYPE
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $ASSIGNMENT_CATEGORY
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = $ORG_GROUP
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate' = ($START_DATE).ToString("dd/MM/yyyy")
    }
    
    Write-Host "Updating user: $AZURE_ID with Job Title: $JOB_TITLE, Department: $DEPARTMENT, Location: $LOCATION, and other attributes."
    
    # Update user Object
    Update-Mguser `
        -UserId $AZURE_ID `
        -JobTitle $JOB_TITLE `
        -Department $DEPARTMENT `
        -EmployeeOrgData $EmployeeOrgData `
        -OfficeLocation $LOCATION `
        -EmployeeHireDate $START_DATE `
        -EmployeeType $POSITION_TYPE `
        -AdditionalProperties $attributes

    Write-Host "Successfully updated user: $AZURE_ID" -ForegroundColor Green
}