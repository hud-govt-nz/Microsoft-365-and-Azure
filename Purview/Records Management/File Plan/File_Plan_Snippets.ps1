Clear-Host
########################################################
# Remove Labels from Purview 

    # Import Label CSV File
    $csvfile = Import-Csv -Path "<CSV File Path>"

    # 1st Removal of Labels
    $csvfile | ForEach-Object {
        $LabelName = $_.LabelName
        Remove-ComplianceTag -Identity $LabelName -Confirm:$false
    }

    # Labels will go into pending deletion state, re-run the command with the force deletion flag to remove fully from Purview.
    Start-Sleep 60
    # 2nd Removal of Labels
    $csvfile | ForEach-Object {
        $LabelName = $_.LabelName
        Remove-ComplianceTag -Identity $LabelName -ForceDeletion -Confirm:$false
    }   

########################################################
# Create File Plan Propoerty Authority in Purview

    # Create Authority
    New-FilePlanPropertyAuthority -Name "GDA6" -Confirm:$false
    # Set meaningful display Name
    Set-FilePlanPropertyAuthority -Identity "GDA6" -DisplayName "General Disposal Authority 6"

########################################################
# Create File Plan Propoerty Citation in Purview

    # Create Citation
    New-FilePlanPropertyCitation -Name "HUD DA" -Confirm:$false
    # Set meaningful display Name
    Set-FilePlanPropertyCitation -Identity "HUD DA" -DisplayName "HUD DA" -citationJurisdiction "HUD"

########################################################
# Create File Plan Propoerty Department in Purview

    # Create Department
    New-FilePlanPropertyDepartment -Name "HUD" -Confirm:$false
    # Set meaningful display Name
    Set-FilePlanPropertyDepartment -Identity "HUD" -DisplayName "HUD"

########################################################
# Create File Plan Propoerty Category in Purview

    # Create Category
    New-FilePlanPropertyCategory -Name "General" -Confirm:$false
    # Set meaningful display Name
    Set-FilePlanPropertyCategory -Identity "General" -DisplayName "General"

########################################################
# Create File Plan Propoerty Subcategory in Purview

    # Create Subcategory
    New-FilePlanPropertySubcategory -Name "General" -Confirm:$false
    # Set meaningful display Name
    Set-FilePlanPropertySubcategory -Identity "General" -DisplayName "General"

########################################################
# referenceId is a custom property that can be used to store a unique identifier for the record.
# This property can be used to link the record to an external system or to a specific record in an external system.

# Create File Plan Propoerty ReferenceId in Purview

    # Create ReferenceId
    New-FilePlanPropertyReferenceId -Name "GDA6" -Confirm:$false

########################################################
