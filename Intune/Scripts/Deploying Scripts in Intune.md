# PowerShell Script Deployment via Microsoft Intune

This document outlines the process of deploying PowerShell scripts through Microsoft Intune. Ensure you have the necessary permissions and prerequisites before proceeding.

## Prerequisites

- An active Microsoft Azure subscription.
- Intune subscription.
- PowerShell script(s) you intend to deploy.

## Getting Started

### 1. Sign in to the Microsoft Endpoint Manager admin center

Navigate to [Endpoint Manager admin center](https://endpoint.microsoft.com/) and sign in with your credentials.

### 2. Navigate to the "Devices" section

Click on `Devices` in the left-hand navigation pane.

### 3. Go to "Scripts"

Under `Devices`, click on `Scripts`.

### 4. Add a new PowerShell script

Click on `+ Add`, and select `PowerShell script`.

### 5. Configure Script Settings

- **Name**: Enter a descriptive name for your script.
- **Description**: Provide a brief description of what the script does.
- **Script Location**: Click on the folder icon and select your PowerShell script.

### 6. Configure Script Settings (Optional)

- **Run this script using the logged-on credentials**: Choose whether the script should run using the logged-on credentials.
- **Enforce script signature check**: Choose whether to enforce a script signature check.
- **Run script in 64-bit PowerShell**: Choose whether to run the script in 64-bit PowerShell.

### 7. Assignments

Assign the script to the desired groups.

### 8. Review + add

Review your configurations and click `Add` to create the script.

### 9. Monitor script deployment

After the script is added, you can monitor its deployment status in the Microsoft Endpoint Manager admin center.

## Troubleshooting

- Ensure you have the necessary permissions to deploy scripts.
- Check the script execution policy settings.
- Review the script status and error messages in the Endpoint Manager admin center.

## Further Reading

- [Microsoft Documentation on PowerShell script deployment](https://docs.microsoft.com/en-us/mem/intune/apps/intune-management-extension)
- [Troubleshooting PowerShell scripts in Intune](https://docs.microsoft.com/en-us/mem/intune/apps/intune-management-extension#troubleshooting-powershell-scripts)

