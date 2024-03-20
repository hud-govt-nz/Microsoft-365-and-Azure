# Connection
Connect-AzAccount -Tenant "9e9b3020-3d38-48a6-9064-373bc7b156dc"

# HUD Azure Subscription
Set-AzContext -Subscription "1c0aeff1-cbf6-4858-9d2b-56645fd92d4c"
New-AzRoleAssignment -ObjectID "02f4690f-6a43-4942-913e-f463cf033b5c" -RoleDefinitionName "Owner" -Scope "/subscriptions/1c0aeff1-cbf6-4858-9d2b-56645fd92d4c" -ObjectType "ForeignGroup"


# HUD-Infrastructure-DEV
Set-AzContext -Subscription "61324dc7-19eb-4999-b8e1-8c7b7ce1cd7f"
New-AzRoleAssignment -ObjectID "02f4690f-6a43-4942-913e-f463cf033b5c" -RoleDefinitionName "Owner" -Scope "/subscriptions/61324dc7-19eb-4999-b8e1-8c7b7ce1cd7f" -ObjectType "ForeignGroup"

# HUD-Infrastructure Prod
Set-AzContext -Subscription "cbc70e9b-2f35-4d2b-a449-abedbb5c7e49"
New-AzRoleAssignment -ObjectID "02f4690f-6a43-4942-913e-f463cf033b5c" -RoleDefinitionName "Owner" -Scope "/subscriptions/cbc70e9b-2f35-4d2b-a449-abedbb5c7e49" -ObjectType "ForeignGroup"

# HUD-Governance
Set-AzContext -Subscription "fa28a6bf-e387-4ae6-a7aa-08366f92d139"
New-AzRoleAssignment -ObjectID "02f4690f-6a43-4942-913e-f463cf033b5c" -RoleDefinitionName "Owner" -Scope "/subscriptions/fa28a6bf-e387-4ae6-a7aa-08366f92d139" -ObjectType "ForeignGroup"

# HUD-Reporting
Set-AzContext -Subscription "3b48a024-bc06-45d3-b011-b75d404aade3"
New-AzRoleAssignment -ObjectID "02f4690f-6a43-4942-913e-f463cf033b5c" -RoleDefinitionName "Owner" -Scope "/subscriptions/3b48a024-bc06-45d3-b011-b75d404aade3" -ObjectType "ForeignGroup" 
