# Connect to Microsoft Graph PowerShell
Write-Host "Connecting to Microsoft Graph"
$Scopes = @('AuditLog.Read.All','Directory.Read.All','Organization.Read.All','User.Read','User.Read.All',"UserAuthenticationMethod.Read.All")
Connect-MgGraph -Scopes $Scopes -NoWelcome | out-null

# Import the CSV file
$Path = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Extract_Phase 1 Azure_Extract_Phase 1 Azure.xlsx"
$Names = Import-Excel -Path $Path

# Loop through each user in Spreadsheet
$Names | ForEach-Object {
    $EMPID = $_.PERSON_NUMBER
    $NAME = $_.NAME
    $UPN = $_.USER_UPN
    $JOBTITLE = $_.JOB_TITLE
    $DEPARTMENT = $_.DEPARTMENT
    $ORGGROUP = $_.GROUPNAME
    $MGRNAME = $_.MGRNAME
    $MGR_UPN = $_.MGR_UPN

    # Assign Manager
    $Managerid = (Get-MgUser -UserId $MGR_UPN).id
    $Manager = @{"@odata.id"="https://graph.microsoft.com/v1.0/users/$Managerid"}
    Set-MgUserManagerByRef -UserId $UPN -BodyParameter $Manager

    # Set Employee ID and OrgData
    $EmployeeOrgData = @{
    'division'  = $ORGGROUP
    }
    
    # Update user Object
    Update-MgBetaUser -UserId $UPN -JobTitle $JOBTITLE -Department $DEPARTMENT -EmployeeId $EMPID -EmployeeOrgData $EmployeeOrgData 


}

# Import the CSV file
$Path = "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\All Staff for Azure_Azure Extract.xlsx"
$Names = Import-Excel -Path $Path

# Loop through each user in Spreadsheet
$Names | ForEach-Object {
    $EMPID = $_.PERSON_NUMBER
    $JOBTITLE = $_.NAME
    $UPN = $_.NAME_UPN
    $DEPARTMENT = $_.DEPARTMENT
    $ORGGROUP = $_.TEAM_GROUP
    $HIREDATE = $_.START_DATE
    $EMPTYPE = $_.POSITION_TYPE
    $EMPCATEGORY = $_.ASSIGNMENT_CATEGORY
   
    $EmpUTCTime = [System.DateTime]::Parse("$HIREDATE")
    $AHOEMPSTARTDATE = $EmpUTCTime.ToUniversalTime().ToLocalTime().ToString("dd/MM/yyyy")
    $EMPHIREDATE = $EmpUTCTime.ToUniversalTime().ToLocalTime()

    # Set Employee ID and OrgData
    $EMPORGDATA = @{
        'division'  = $ORGGROUP
        }

    $attributes = @{
	    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType' = $EMPTYPE
	    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory' = $EMPCATEGORY
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = $ORGGROUP
        'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate' = $AHOEMPSTARTDATE
        }

    <# Clear Employee ID
    try {
        Invoke-GraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/$UPN" -Body '{"employeeId": null}'
        Write-Host "Updated $UPN successfully" -ForegroundColor Green
        } catch {
                $ErrorMessage = $_.Exception.Message
            Write-Host "Error updating $UPN" -ForegroundColor Red
            Write-Host $ErrorMessage
        }
        #>
    Try{
        # Update user Object
        Update-MgBetaUser `
            -UserId $UPN `
            -EmployeeOrgData $EMPORGDATA `
            -EmployeeHireDate $EMPHIREDATE `
            -EmployeeType $EMPTYPE `
            -AdditionalProperties $attributes

        Write-Host "Updated $UPN successfully" -ForegroundColor Green
    }
    Catch{
        $ErrorMessage = $_.Exception.Message
        Write-Host "Error updating $UPN" -ForegroundColor Red
        Write-Host $ErrorMessage
    }


}