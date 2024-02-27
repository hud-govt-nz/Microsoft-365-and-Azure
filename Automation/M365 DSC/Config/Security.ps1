# Getting application information for Application + Certificate authentication
$ApplicationId = "c574d912-050b-4acc-b0e6-63bed2c4c562"
$CertificateThumbprint = "789C2920A65EE8D5F4C3DDF4487B46B6D29293E4"
$TenantId = "mhud.onmicrosoft.com"
$Path =".\Security"

# Define workload components
$Components = @(
    "SCAuditConfigurationPolicy",
    "SCAutoSensitivityLabelPolicy",
    "SCAutoSensitivityLabelRule",
    "SCCaseHoldPolicy",
    "SCCaseHoldRule",
    "SCComplianceCase",
    "SCComplianceSearch",
    "SCComplianceSearchAction",
    "SCComplianceTag",
    "SCDeviceConditionalAccessPolicy",
    "SCDeviceConfigurationPolicy",
    "SCDLPCompliancePolicy",
    "SCDLPComplianceRule",
    "SCFilePlanPropertyAuthority",
    "SCFilePlanPropertyCategory",
    "SCFilePlanPropertyCitation",
    "SCFilePlanPropertyDepartment",
    "SCFilePlanPropertyReferenceId",
    "SCFilePlanPropertySubCategory",
    "SCLabelPolicy",
    "SCProtectionAlert",
    "SCRetentionCompliancePolicy",
    "SCRetentionComplianceRule",
    "SCRetentionEventType",
    "SCSecurityFilter",
    "SCSensitivityLabel",
    "SCSupervisoryReviewPolicy",
    "SCSupervisoryReviewRule"
)

# Get current date
$CurrentDate = (Get-Date).ToString('MMddyy')

# Construct the file name and configuration name
$FileName = "M365DSCConfig_$CurrentDate.ps1"
$ConfigName = "M365DSCConfig_$CurrentDate"

# Exporting resources using certificate
Export-M365DSCConfiguration `
    -GenerateInfo $true `
    -MaxProcesses 4 `
    -Path $Path `
    -FileName $FileName `
    -ConfigurationName $ConfigName `
    -TenantId $TenantID `
    -ApplicationId $ApplicationID `
    -CertificateThumbprint $CertificateThumbprint `
    -Components $Components
