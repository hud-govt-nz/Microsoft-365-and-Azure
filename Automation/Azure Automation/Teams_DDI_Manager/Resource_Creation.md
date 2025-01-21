# Automating Teams Direct-Dial-In (DDI) Assignment using Azure Automation

The following is a step-by-step guide on how to create an Azure Automation account with all the necessary dependancies which can then be triggered via a webhook. Specifically, this automation will manage the assignment of DDI connections to users within an MS Teams environment and update the details in Entra. 

The webhook in this use case is going to be used within Dynamics so that whenever a number is updated/assigned in that system it makes the necessary changes within M365. 

## Outline 

### 1. Azure and Service Account Creation  
  - Includes creation of
    - Azure Resource Group
    - Azure Automation Account
      - Assignment of required API permissions
    - Entra User account
      - Assignment of required administrative roles

### 2. Setting up Automation dependancies
  - Setting up Entra user details as shared resource on automation account
  - Installing necessary PowerShell Modules

### 3. Runbook creation
  - Setup of runbooks to:
    - connection to Microsoft Graph using PowerShell SDK
    - connection to MicrosoftTeams 
    - Assigning DDI number to end user
  
### 4. Webhook setup for integration with Dynamics 365
  - Enabling webhook


## 1. Azure and Service Account Creation

Below is the steps using the Az and Graph PowerShell modules to create the necessary resources for this automation to work. Note that all of this can be done via the web portal **except** the assignment of API permissions. 
[Grant Graph API Permission to Managed Identity Object](https://techcommunity.microsoft.com/t5/azure-integration-services-blog/grant-graph-api-permission-to-managed-identity-object/ba-p/2792127)

  
```powershell  
# Create Azure Resource Group, Automation Account, and Managed Identity  
Import-Module Az  
Connect-AzAccount  
  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  
  
$resourceGroupName = "syd-rg-app-teams-ddi-management"  
$location = "Australia East"  
$automationAccountName = "TeamsDDIManager" # Automation Account  
  
# Create the Resource Group  
New-AzResourceGroup -Name $resourceGroupName -Location $location  
  
# Create the Automation Account within the specified Resource Group  
New-AzAutomationAccount -ResourceGroupName $resourceGroupName -Name $automationAccountName -Location $location  
  
# Enable System Managed Identity on the Automation Account  
$automationAccount = Get-AzAutomationAccount -ResourceGroupName $resourceGroupName -Name $automationAccountName  
$automationAccount | Set-AzAutomationAccount -AssignSystemIdentity $true  
  
# Assign Graph permissions to Managed Identity  
Import-Module Microsoft.Graph   
Connect-MgGraph # can define scope here but since we are using this same session to also create the new user you can just connect generally.  
  
# Get GUID of Managed Identity  
$managedIdentityId = (Get-MgServicePrincipal -Filter "displayName eq 'TeamsDDIManager'").id  
  
# Assign Graph Permissions to Managed Identity  
$graphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'" #AppId of Microsoft Graph in all Enterprise Applications, always the same in each tenant.  
$graphScopes =@('User.ReadWrite.All','Directory.Read.All','Organization.Read.All')  
  
ForEach($scope in $graphScopes){  
    $appRole = $graphApp.AppRoles | Where-Object {$_.Value -eq $scope}  
    New-MgServicePrincipalAppRoleAssignment -PrincipalId $managedIdentityId -ServicePrincipalId $managedIdentityId -ResourceId $graphApp.Id -AppRoleId $appRole.Id  
}  
  
# Assign Skype Tenant Admin API Permissions to Managed Identity  
  
$SkypeAPI =@('application_access_custom_sba_appliance','application_access')  
$SkypeAPI = Get-MgServicePrincipal -Filter "AppId eq '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'" #AppId of Skype and Teams Tenant Admin API in all Enterprise Applications, always the same in each tenant.  
  
ForEach($scope in $SkypeAPI){  
    $appRole = $SkypeAPI.AppRoles | Where-Object {$_.Value -eq $scope}  
    New-MgServicePrincipalAppRoleAssignment -PrincipalId $managedIdentityId -ServicePrincipalId $managedIdentityId -ResourceId $SkypeAPI.Id -AppRoleId $appRole.Id  
}  
  
# Create service account and assign Team Administrator role  
  
# New User  
$PasswordProfile = @{}  
    $PasswordProfile["Password"]= "ComplexPassword!123"  # Production Password Stored in BitWarden
    $PasswordProfile["ForceChangePasswordNextSignIn"] = $true  
  
New-MgUser `  
    -AccountEnabled $true `  
    -DisplayName "SVC_Teams_DDI_Manager" `  
    -MailNickname "SVC_Teams_DDI_Manager" `  
    -UserPrincipalName "SVC_Teams_DDI_Manager@hud.govt.nz" `  
    -PasswordPolicies DisablePasswordExpiration  
    -PasswordProfile $PasswordProfile  
  
# Get the user's Object ID  
$userId = (Get-MgUser -UserId "SVC_Teams_DDI_Manager@hud.govt.nz").Id  
  
# Get the Teams Administrator Directory Role Object ID  
$directoryRoleId = (Get-MgDirectoryRole -Filter "displayName eq 'Teams Administrator'").Id  
  
# Assign the Teams Administrator role to the new user  
New-MgDirectoryRoleMemberByRef -DirectoryRoleId "$directoryRoleId" -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($userId)"}  

```

## 2. Setting up Automation dependancies

### 2.1 Adding Additional PowerShell modules

Before any runbooks are created in the automation account it needs to have the necessary PS modules installed in order to run Teams and Graph commands, follow these steps:

You can action this via the portal or connect to the automation account using Azure PowerShell again. [Learn about Shared Resources and Modules in Azure Automation](https://learn.microsoft.com/en-us/azure/automation/shared-resources/modules)


``` PowerShell
Import-Module Az  
Connect-AzAccount  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  

# Define the name of the Resource Group and the Automation Account
$resourceGroupName = "syd-rg-app-teams-ddi-management"  
$automationAccountName = "TeamsDDIManager" # Automation Account  

# List of module names to be imported
$moduleNames = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'MicrosoftTeams')

# Import each module into the Automation Account from the PowerShell Gallery
foreach ($moduleName in $moduleNames) {
    New-AzAutomationModule -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Name $moduleName -ContentLinkUri "https://www.powershellgallery.com/packages/$moduleName"
}
```
### 2.2 Adding a Service Account User Account as a Credential in Azure Automation Account

Follow the steps below to add a new user account as a set of credentials on an Azure Automation Account via the Azure Portal.


  #### 1. **Login to Azure Portal**:
  - Navigate to the [Azure Portal](https://portal.azure.com/) and log in with your credentials.

  #### 2. **Access Your Automation Account**:
  - In the left-hand menu, select "All services" then find and select "Automation Accounts" from the list.
  - Click on the name of your Automation Account from the list.

  #### 3. **Create a Credential Asset**:
  - Under the section "Shared Resources", select "Credentials".
  - Click on "+ Add a credential" at the top of the pane.
  - In the "Add Credential" pane, provide the following details:
    - **Name**: Enter a name for the credential (e.g., MyUserCredential).
    - **User name**: Enter the username of the new user account you created (e.g., JohnD@example.com).
    - **Password** and **Confirm password**: Enter the password of the new user account.
  - Click on "Create" to save the credential.

  #### 4. **Verification**:
  - Once the credential asset is created, it will be listed under "Credentials" in your Automation Account. You can click on the credential name to view its details (though the password will not be visible for security reasons).

  Now you have successfully added a new user account as a set of credentials on your Azure Automation Account via the Azure Portal.
 
### 2.3 Alternatively done via the shell

 ```Powershell
Import-Module Az  
Connect-AzAccount  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  

# Define the name of the Resource Group, Automation Account, and the credentials
$resourceGroupName = "syd-rg-app-teams-ddi-management"  
$automationAccountName = "TeamsDDIManager" # Automation Account

# New shared credential resource for automation account
$credentialName = "TeamsAdminAccount"
$userName = "SVC_Teams_DDI_Manager@hud.govt.nz"  # Replace with the user principal name of the user you created
$password = ConvertTo-SecureString "ComplexPassword!123" -AsPlainText -Force  # Replace with the password of the user you created

# Create the credential asset in the Automation Account
New-AzAutomationCredential -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Name $credentialName -Value (New-Object System.Management.Automation.PSCredential ($userName, $password))
```
## 3. Runbook creation

An Azure Automation runbook is a set of predefined tasks or operations that can be executed to manage and orchestrate processes in your Azure environment. They can be authored in various scripting languages including PowerShell, Python, and Graph, allowing for automation of routine tasks, complex deployments, and even automatic troubleshooting and remediation. [Learn about Azure Automation Runbook Types](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-types?tabs=lps51%2Cpy27)

In the example below a single runbook is created named **MyRunbook**. In the production runbook the actions for connecting to Graph, Teams, and the DDI assignment have been separated to make them easier to manage.

 ```Powershell
Import-Module Az  
Connect-AzAccount  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  

# Define the name of the Resource Group, Automation Account, and the credentials
$resourceGroupName = "syd-rg-app-teams-ddi-management"  
$automationAccountName = "TeamsDDIManager" # Automation Account
$runbookName = "MyRunbook"
$runbookType = "PowerShell"  # Other types include: Graph, Python, PowerShellWorkflow

# Create the new runbook
New-AzAutomationRunbook -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -Name $runbookName -Type $runbookType
```

### 3.1 screenshot of production
![Runbooks](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/main/Azure%20Automation/Teams_DDI_Manager/Artifacts/runbooks.png?raw=true)

### 3.2 Runbooks

[Graphlogin.ps1 Script](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/4541cbdd529725eaefcb83a26b111ea0dfc291ca/Azure%20Automation/Teams_DDI_Manager/Runbooks/Graphlogin.ps1)

[TeamsLogin.ps1 Script](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/4541cbdd529725eaefcb83a26b111ea0dfc291ca/Azure%20Automation/Teams_DDI_Manager/Runbooks/TeamsLogin.ps1)

[Update_Teams_DDI.ps1 Script](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/4541cbdd529725eaefcb83a26b111ea0dfc291ca/Azure%20Automation/Teams_DDI_Manager/Runbooks/Update_Teams_DDI.ps1)

## 4 Webhook creation
To create a webhook run the following code snippet
``` PowerShell

Import-Module Az  
Connect-AzAccount  
Set-AzContext -SubscriptionId "Enter HUD Azure Subscription ID"  

# Define the name of the Resource Group, Automation Account, and the credentials
$resourceGroupName = "syd-rg-app-teams-ddi-management"  
$automationAccountName = "TeamsDDIManager" # Automation Account
$runbookName = "Update_Teams_DDI"
$webhookName = "Update_DDI"

# Create the new webhook
New-AzAutomationWebhook -ResourceGroupName $resourceGroupName -AutomationAccountName $automationAccountName -RunbookName $runbookName -Name $webhookName -IsEnabled $true -ExpiryTime (Get-Date).AddYears(2)

```

### 4.1 Screenshot of Production

![Webhooks](https://github.com/hud-govt-nz/Microsoft-365-and-Azure/blob/main/Azure%20Automation/Teams_DDI_Manager/Artifacts/Webhooks.png?raw=true)

### 4.2 How it works
The webhook is used to activate a process when triggered. In this instance when triggered it will run the Runbook Update_Teams_DDI

The event trigger also sends through the necessary information that needs to be updated, in this instance that would be the users email address and the DDI number being assigned/updated. This needs to be expressed in JSON format to work correctly.

``` json
//Example JSON formatting
{
  "User": "UserPrincipalName",
  "DDI": "No in +64 4 *** ****"
}
```
The runbook Update_Teams_DDI will pick this informaiton up and parse it through to a useable format that can be used by powershell. 

The snippet below is taken from the beginning of the Update_Teams_DDI runbook and is used to capture and transform the json data into a usable format in PS.

``` Powershell
# Define parameter block
Param
(
    [Parameter (Mandatory = $true)]
    [object] $WebhookData
)

...

# Extracting parameters from WebhookData
$WebhookBody = $WebhookData.RequestBody | ConvertFrom-Json
$User = $WebhookBody.User
$DDI = $WebhookBody.DDI

```