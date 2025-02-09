Clear-Host

# Retrieve all compliance tags
$complianceTags = Get-ComplianceTag

# Initialize an array to store the processed data
$processedData = @()

# Loop through each compliance tag and expand the FilePlanMetadata
foreach ($tag in $complianceTags) {
    $metadata = $tag.FilePlanMetadata | ConvertFrom-Json
    $properties = @{}

    foreach ($setting in $metadata.Settings) {
        $properties[$setting.Key] = $setting.Value
    }

    # Add the tag name, expanded properties, and additional properties to the processed data
    $processedData += [PSCustomObject]@{
        Name = $tag.Name
        Comment = $tag.Comment
        Notes = $tag.Notes
        IsRecordLabel = $tag.IsRecordLabel
        RetentionAction = $tag.RetentionAction
        RetentionDuration = $tag.RetentionDuration
        RetentionType = $tag.RetentionType
        ReviewerEmail = $tag.ReviewerEmail
        Regulatory = $tag.Regulatory
        FilePlanPropertyDepartment = $properties["FilePlanPropertyDepartment"]
        FilePlanPropertyCategory = $properties["FilePlanPropertyCategory"]
        FilePlanPropertySubcategory = $properties["FilePlanPropertySubcategory"]
        FilePlanPropertyCitation = $properties["FilePlanPropertyCitation"]
        FilePlanPropertyReferenceId = $properties["FilePlanPropertyReferenceId"]
        FilePlanPropertyAuthority = $properties["FilePlanPropertyAuthority"]
    }
}

# Export the processed data to a CSV file
$processedData | Export-Csv -Path "<CSV File Path>" -NoTypeInformation
