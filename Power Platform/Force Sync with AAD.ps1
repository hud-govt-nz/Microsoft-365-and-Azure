#onnect to PowerApps and AzureAD modules
Connect-AzureAD

#Obtain list of Environments in PowerApps
Get-AdminPowerAppEnvironment | Select-Object DisplayName, EnvironmentName

#Specify Environments to sync
$Env =@{
    0 = "ba402bb3-6e51-e30e-964c-b5c4b22c48cd" #Customer Service Dev
    1 = "7397bd0c-d335-e820-90be-ee5d5c4c61c4" #Customer Service UAT
    2 = "ae59b500-5f9b-e129-9a63-582bb44ec6c2" #Customer Service Prod
}

#Select user to manually sync
$User = Get-AzureADUser -ObjectId "Dan.Michaels@hud.govt.nz"

# f0fefe16-169d-469b-81b6-03171720d228 = User - All HUD Staff
Add-AzureADGroupMember -ObjectId "f0fefe16-169d-469b-81b6-03171720d228" -RefObjectId $User.objectid

#Sync to each PP environment
Foreach ($Key in $Env.Keys) {
    Add-AdminPowerAppsSyncUser -EnvironmentName $($Env[$Key]) -PrincipalObjectId $User.objectid -Verbose

}

#Get-AzureADUserExtension -ObjectId $User.ObjectId
Get-AzureADUserExtension -ObjectId $User.ObjectId

