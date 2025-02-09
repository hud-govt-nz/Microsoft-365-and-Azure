Clear-Host

# Import Label CSV File
$csvfile = Import-Csv -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Purview Tidy Up\Records Management\CSV Import Files\Record_Label_File_Plan_Import.csv"

# Create Label
$csvfile | ForEach-Object {

    # Label Properties
    $LabelName = $_.LabelName
    $Notes = $_.Notes
    $RecordLabel = [bool]::Parse($_.IsRecordLabel)
    $Action = $_.RetentionAction
    $Duration = $_.RetentionDuration
    $Type = $_.RetentionType
    $Regulatory = [bool]::Parse($_.Regulatory)
    
    # File plan properties
    $RefID = $_.ReferenceId
    $Department = $_.DepartmentName
    $Category = $_.Category
    $SubCategory = $_.SubCategory
    $Citation = $_.CitationName
    $Authority = $_.AuthorityType

    # Multi Stage Review Properties
    $StageName = $_.ReviewStageName
    $ReviewEmail = $_.ReviewerEmail
    
    # Place FilePlanProperties into custom object.
    # https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancetag?view=exchange-ps#-fileplanproperty
    $FilePlanData = [PSCustomObject]@{  
        Settings = @(
            @{Key = "FilePlanPropertyReferenceId"; Value = $RefID},
            @{Key = "FilePlanPropertyDepartment"; Value = $Department},
            @{Key = "FilePlanPropertyCategory"; Value = $Category},
            @{Key = "FilePlanPropertySubcategory"; Value = $SubCategory},  
            @{Key = "FilePlanPropertyCitation"; Value = $Citation},
            @{Key = "FilePlanPropertyAuthority"; Value = $Authority}               
        )  
    }  

    # Transform FilePlanProperties into JSON object
    # https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancetag?view=exchange-ps#-fileplanproperty
    $FilePlan = ConvertTo-Json $FilePlanData
    
    # Define the single reviewer for the stage
    # https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancetag?view=exchange-ps#-multistagereviewproperty
    $MultiStageReviewProperty = '{"MultiStageReviewSettings":[{"StageName":"' + $StageName + '","Reviewers":["' + $ReviewEmail + '"]}]}'

    #Create Labels
    try {
        if ($Action -eq "KeepAndDelete") {
            # Create a new compliance tag  
            $Tags = New-ComplianceTag `
                        -Name $LabelName `
                        -Notes $Notes `
                        -IsRecordLabel $RecordLabel `
                        -RetentionAction $Action `
                        -RetentionDuration $Duration `
                        -RetentionType $Type `
                        -FilePlanProperty $FilePlan `
                        -Regulatory $Regulatory `
                        -MultiStageReviewProperty $MultiStageReviewProperty
            Write-Host "Label $($Tags.Name) has been created." -ForegroundColor Green
        } else {
            # Create a new compliance tag excluding disposition review.
            # https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancetag?view=exchange-ps#-revieweremail 
            $Tags = New-ComplianceTag `
                        -Name $LabelName `
                        -Notes $Notes `
                        -IsRecordLabel $RecordLabel `
                        -RetentionAction $Action `
                        -RetentionDuration $Duration `
                        -RetentionType $Type `
                        -FilePlanProperty $FilePlan `
                        -Regulatory $Regulatory
            Write-Host "Label $($Tags.Name) has been created." -ForegroundColor Green
        }
    } catch {
            # Display the error message
            Write-Output "An error occurred: $_"
    }    
}