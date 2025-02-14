## CREATE LABEL POLICY ##

    # Import Policy CSV File
    # CSV can only have one publishedtag (label) against it or the rule cmdlet fails. Have to add additional labels to policy in Purview Portal.
    $Policies = Import-Csv -Path "C:\Users\Ashley.Forde\OneDrive - Ministry of Housing and Urban Development\Desktop\Purview Tidy Up\Records Management\CSV Import Files\Record_Label_Policy_Import.csv"

    # Create Policies
    $Policies | ForEach-Object {
        $PolicyName = $_.PolicyName
        $Comment = $_.Comment
        $Tags = $_.PublishComplianceTag
        $SharePointLocation = $_.SharePointLocation
        $Policy = New-RetentionCompliancePolicy -Name $PolicyName -Comment $Comment -SharePointLocation $SharePointLocation -Enabled:$false
        New-RetentionComplianceRule -Policy $Policy.guid -PublishComplianceTag $Tags 
    }



    Get-retentioncompliancepolicy
    

    