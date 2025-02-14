Clear-Host

########################################################
# Remove Labels from Purview 

Write-Host "Importing Label CSV file..."
# Import Label CSV File
$csvfile = Import-Csv -Path "<CSV File Path>"

Write-Host "Starting 1st Removal of Labels..."
# 1st Removal of Labels
$csvfile | ForEach-Object {
    $LabelName = $_.LabelName
    Write-Host "Removing label: $LabelName"
    Remove-ComplianceTag -Identity $LabelName -Confirm:$false
}

Write-Host "Waiting for labels to go into pending deletion state..."
# Labels will go into pending deletion state, re-run the command with the force deletion flag to remove fully from Purview.
Start-Sleep 60

Write-Host "Starting force deletion of Labels..."
# 2nd Removal of Labels
$csvfile | ForEach-Object {
    $LabelName = $_.LabelName
    Write-Host "Force deleting label: $LabelName"
    Remove-ComplianceTag -Identity $LabelName -ForceDeletion -Confirm:$false
}

########################################################
# Create File Plan Property Authority in Purview

Write-Host "Creating File Plan Property Authority 'GDA6'..."
# Create Authority
New-FilePlanPropertyAuthority -Name "GDA6" -Confirm:$false

Write-Host "Setting display name for Authority 'GDA6'..."
# Set meaningful display Name
Set-FilePlanPropertyAuthority -Identity "GDA6" -DisplayName "General Disposal Authority 6"

########################################################
# Create File Plan Property Citation in Purview

Write-Host "Creating File Plan Property Citation 'HUD DA'..."
# Create Citation
New-FilePlanPropertyCitation -Name "HUD DA" -Confirm:$false

Write-Host "Setting display name for Citation 'HUD DA'..."
# Set meaningful display Name
Set-FilePlanPropertyCitation -Identity "HUD DA" -DisplayName "HUD DA" -citationJurisdiction "HUD"

########################################################
# Create File Plan Property Department in Purview

Write-Host "Creating File Plan Property Department 'HUD'..."
# Create Department
New-FilePlanPropertyDepartment -Name "HUD" -Confirm:$false

Write-Host "Setting display name for Department 'HUD'..."
# Set meaningful display Name
Set-FilePlanPropertyDepartment -Identity "HUD" -DisplayName "HUD"

########################################################
# Create File Plan Property Category in Purview

Write-Host "Creating File Plan Property Category 'General'..."
# Create Category
New-FilePlanPropertyCategory -Name "General" -Confirm:$false

Write-Host "Setting display name for Category 'General'..."
# Set meaningful display Name
Set-FilePlanPropertyCategory -Identity "General" -DisplayName "General"

########################################################
# Create File Plan Property Subcategory in Purview

Write-Host "Creating File Plan Property Subcategory 'General'..."
# Create Subcategory
New-FilePlanPropertySubcategory -Name "General" -Confirm:$false

Write-Host "Setting display name for Subcategory 'General'..."
# Set meaningful display Name
Set-FilePlanPropertySubcategory -Identity "General" -DisplayName "General"

########################################################
# Create File Plan Property ReferenceId in Purview

Write-Host "Creating File Plan Property ReferenceId 'GDA6'..."
# Create ReferenceId
New-FilePlanPropertyReferenceId -Name "GDA6" -Confirm:$false