# Generate new attributes for Aho Azure Integration
# Connect to Graph
$Scopes =@('Application.ReadWrite.All','Application.ReadWrite.OwnedBy','Directory.ReadWrite.All','User.ReadWrite.All')
Connect-MgGraph -Scopes $Scopes -NoWelcome | Out-Null


# Get Existing Application Details
$AppObjectID = "be7a8282-b15f-48a3-a5f5-0b59be13f0f8"

Get-MgApplicationExtensionProperty -ApplicationId $AppObjectID

# Add Attribute to extension properties
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.applications/new-mgapplicationextensionproperty?view=graph-powershell-1.0

$params = @{
	name = "<NAMEOFVALUEHERE>"
	dataType = "String"
	isMultiValued = $false
	targetObjects = @(
		"User"
		"Group"
	)
}

New-MgApplicationExtensionProperty -ApplicationId $applicationId -BodyParameter $params

# Example Add attribute to user
$User = Get-MgBetaUser -UserId "Ashley.Forde@hud.govt.nz"


$Attributes = @{
    'extension_56a473fa1d5b476484f306f7b06ee688_ObjectUserOrganisationalGroup' = 'OrgGroup_Test2'
    }

Update-MgBetaUser -UserId $User.id -AdditionalProperties $Attributes