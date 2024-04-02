# Connect to Microsoft Graph PowerShell
Write-Host "Connecting to Microsoft Graph"
$Scopes = @('AuditLog.Read.All','Directory.Read.All','Organization.Read.All','User.Read','User.Read.All')
Connect-MgGraph -Scopes $Scopes -NoWelcome | out-null

# Obtain Last Sign Date Time (Non Standard property value)
Write-Host "Collecting user information..." -ForegroundColor yellow
$Results = @()
$results += Get-MgBetaUser -All -Property id,SignInActivity | Select-Object -Property id,@{ Name = 'LastSignInDateTime'; Expression = { [datetime]$_.SignInActivity.LastSignInDateTime } }

# Gather other Attributes.
$Values = @()

$Users = Get-MgBetaUser -All | Select-Object ID,CreatedDateTime,AccountEnabled,UserType,DisplayName,GivenName,Surname,UserPrincipalName,Mail,UsageLocation,
Department,JobTitle,CompanyName,StreetAddress,City,PostalCode,State,Country,officelocation,SecurityIdentifier,MobilePhone,
@{ Name = 'BusinessPhones'; Expression = { [string]$_.BusinessPhones -replace "{",'' } },
@{ Name = 'passwordPolicies'; Expression = { [string]$_.passwordPolicies } },
@{ Name = "StartDate"; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserStartDate'] } },
@{ Name = "LeaveDate"; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserLeaveDateTime'] } },
@{ Name = "EmployeeType"; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmployeeType'] } },
@{ Name = "EmployeeCategory"; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserEmploymentCategory'] } },
@{ Name = "OrgGroup"; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup'] } },
@{ Name = 'M365E3'; Expression = { if ($_.assignedLicenses.skuid -eq "05e9a617-0261-4cee-bb44-138d3ef5d965") { $true } else { $false } } },
@{ Name = 'M365E5'; Expression = { if ($_.assignedLicenses.skuid -eq "06ebc4ee-1bb5-47dd-8120-11324bc54e06") { $true } else { $false } } },
@{ Name = 'NoLicense'; Expression = { ($_.assignedLicenses.count -eq 0) } },
@{ Name = 'RoomMailbox'; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_RoomMailbox'] } },
@{ Name = 'SharedMailbox'; Expression = { $_.AdditionalProperties['extension_56a473fa1d5b476484f306f7b06ee688_SharedMailbox'] } }

foreach ($User in $Users) {
    try {
        $Manager = Get-MgUserManager -UserId $User.ID -ErrorAction SilentlyContinue
        $User | Add-Member -MemberType NoteProperty -Name Manager -Value $Manager.AdditionalProperties.userPrincipalName -ErrorAction SilentlyContinue
        $Values += $User
    } catch [Microsoft.Graph.ServiceException] {
        if ($_.Exception.Error.Code -eq "Request_ResourceNotFound") {
            Write-Host "Manager not found for user: $($User.ID)"
        } else {
            Write-Host "Error retrieving manager for user: $($User.ID)"
            Write-Host "Error message: $($_.Exception.Message)"
        }
        continue
    }
}
# Assuming $Results and $Values are the two arrays

# Merge the arrays
for ($i = 0; $i -lt $Results.count; $i++) {
	$result = $Results[$i]
	$value = $Values[$i]

	$result | Add-Member -MemberType NoteProperty -Name 'CreatedDateTime' -Value $value.CreatedDateTime
	$result | Add-Member -MemberType NoteProperty -Name 'AccountEnabled' -Value $value.accountenabled
	$result | Add-Member -MemberType NoteProperty -Name 'UserType' -Value $value.UserType
	$result | Add-Member -MemberType NoteProperty -Name 'DisplayName' -Value $value.DisplayName
	#$result | Add-Member -MemberType NoteProperty -Name 'GivenName' -Value $value.GivenName
	#$result | Add-Member -MemberType NoteProperty -Name 'Surname' -Value $value.Surname
	$result | Add-Member -MemberType NoteProperty -Name 'UserPrincipalName' -Value $value.UserPrincipalname
	$result | Add-Member -MemberType NoteProperty -Name 'Mail' -Value $value.Mail
	$result | Add-Member -MemberType NoteProperty -Name 'UsageLocation' -Value $value.UsageLocation
	$result | Add-Member -MemberType NoteProperty -Name 'Department' -Value $value.Department
	$result | Add-Member -MemberType NoteProperty -Name 'JobTitle' -Value $value.Jobtitle
	$result | Add-Member -MemberType NoteProperty -Name 'CompanyName' -Value $value.CompanyName
	$result | Add-Member -MemberType NoteProperty -Name 'StreetAddress' -Value $value.StreetAddress
	$result | Add-Member -MemberType NoteProperty -Name 'City' -Value $value.City
	$result | Add-Member -MemberType NoteProperty -Name 'PostalCode' -Value $value.PostalCode
	$result | Add-Member -MemberType NoteProperty -Name 'State' -Value $value.State
	$result | Add-Member -MemberType NoteProperty -Name 'Country' -Value $value.Country
	$result | Add-Member -MemberType NoteProperty -Name 'OfficeLocation' -Value $value.officelocation
	#$result | Add-Member -MemberType NoteProperty -Name 'SecurityIdentifier' -Value $value.SecurityIdentifier
	$result | Add-Member -MemberType NoteProperty -Name 'MobilePhone' -Value $value.MobilePhone
	$result | Add-Member -MemberType NoteProperty -Name 'BusinessPhones' -Value $value.BusinessPhones
	$result | Add-Member -MemberType NoteProperty -Name 'passwordPolicies' -Value $value.passwordPolicies
	$result | Add-Member -MemberType NoteProperty -Name 'StartDate' -Value $value.StartDate
	$result | Add-Member -MemberType NoteProperty -Name 'LeaveDate' -Value $value.LeaveDate
	$result | Add-Member -MemberType NoteProperty -Name 'EmployeeType' -Value $value.EmployeeType
	$result | Add-Member -MemberType NoteProperty -Name 'EmployeeCategory' -Value $value.EmployeeCategory
	$result | Add-Member -MemberType NoteProperty -Name 'OrgGroup' -Value $value.OrgGroup
	$result | Add-Member -MemberType NoteProperty -Name 'M365E3' -Value $value.M365E3
	$result | Add-Member -MemberType NoteProperty -Name 'M365E5' -Value $value.M365E5
	$result | Add-Member -MemberType NoteProperty -Name 'NoLicense' -Value $value.NoLicense
	$result | Add-Member -MemberType NoteProperty -Name 'RoomMailbox' -Value $value.RoomMailbox
	$result | Add-Member -MemberType NoteProperty -Name 'SharedMailbox' -Value $value.SharedMailbox
	$result | Add-Member -MemberType NoteProperty -Name 'Manager' -Value $value.Manager
	
}

$Output = $results | Select-Object `
	ID, `
	DisplayName, `
	JobTitle, `
	Department, `
	@{Name='Organisational Group (Aho)'; Expression={$_.'OrgGroup'}}, `
	CompanyName, `
	Mail,`
	UserPrincipalName, `
	BusinessPhones, `
	MobilePhone, `
	manager, `
	StreetAddress, `
	City, `
	PostalCode, `
	State, `
	Country, `
	officelocation, `
	AccountEnabled, `
	UserType, `
	CreatedDateTime, `
	@{Name='StartDate (Aho)'; Expression={$_.'StartDate'}}, `
	@{Name='EmployeeType (Aho)'; Expression={$_.'EmployeeType'}}, `
	@{Name='EmployeeCategory (Aho)'; Expression={$_.'EmployeeCategory'}}, `
	M365E3, `
	M365E5, `
	NoLicense, `
	UsageLocation, `
	passwordPolicies, `
	RoomMailbox, `
	SharedMailbox, `
	LastSignInDateTime,`
	@{Name='Leave Date (Aho)'; Expression={$_.'LeaveDate'}} | Sort-Object DisplayName

Write-Host "Open Save Dialog"

$Date = Get-Date -Format "dd.MM.yyyy h.mm tt"
$FileName = "Entra All Users Export"

# Add assembly and import namespace  
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog

# Configure the SaveFileDialog  
$SaveFileDialog.Filter = "XLSX Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$SaveFileDialog.Title = "Save as"
$SaveFileDialog.FileName = $FileName

# Show the SaveFileDialog and get the selected file path  
$SaveFileResult = $SaveFileDialog.ShowDialog()
if ($SaveFileResult -eq [System.Windows.Forms.DialogResult]::OK) {
	$SelectedFilePath = $SaveFileDialog.FileName
	$Output | Export-Excel $SelectedFilePath -AutoSize -AutoFilter -WorksheetName $Date -FreezeTopRow -BoldTopRow
    
    $excelPackage = Open-ExcelPackage -Path $SelectedFilePath
    $worksheet = $excelPackage.Workbook.Worksheets["$Date"]

    # Assuming headers are in row 1 and you start from row 2
    $startRow = 1
    $endRow = $worksheet.Dimension.End.Row
    $startColumn = 1
    $endColumn = $worksheet.Dimension.End.Column

    # Set horizontal alignment to left for all cells in the used range
    for ($col = $startColumn; $col -le $endColumn; $col++) {
        for ($row = $startRow; $row -le $endRow; $row++) {
            $worksheet.Cells[$row, $col].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
        }
    }

   
    # Autosize columns if needed
    foreach ($column in $worksheet.Dimension.Start.Column..$worksheet.Dimension.End.Column) {
        $worksheet.Column($column).AutoFit()
    }
    
    # Save and close the Excel package
    $excelPackage.Save()
    Close-ExcelPackage $excelPackage -Show

	Write-Host "The report $FileName has been saved in $($SelectedFilePath)" -ForegroundColor Green
} else {
	Write-Host "Save cancelled" -ForegroundColor Yellow
}
