# Creating a .intunewin File for Deploying Win32 Apps from Intune  
   
This guide will walk you through the process of creating a .intunewin file for deploying Win32 apps through Microsoft Intune.  
   
## Prerequisites  
   
1. Download and install the [Microsoft Win32 Content Prep Tool](https://github.com/Microsoft/Microsoft-Win32-Content-Prep-Tool) from GitHub.  
2. Ensure you have the installer file for the Win32 app you want to deploy (e.g., .msi or .exe or PS script).  
   
## Steps to Create a .intunewin File  
   
1. Open a command prompt with administrative privileges.  
2. Navigate to the directory where the Microsoft Win32 Content Prep Tool is installed, e.g.:  
  
   ```  
   cd C:\Win32ContentPrepTool  
   ```  
   
3. Run the following command to create a .intunewin file:  
  
   ```  
   IntuneWinAppUtil.exe -c <source_folder> -s <setup_file> -o <output_folder>  
   ```  
  
   Replace the following placeholders with appropriate values:  
   - `<source_folder>`: The path to the folder containing the installer file for your Win32 app (e.g., `C:\Win32Apps\MyApp`).  
   - `<setup_file>`: The name of the installer file for your Win32 app (e.g., `MyApp.msi`).  
   - `<output_folder>`: The path to the folder where you want the .intunewin file to be saved (e.g., `C:\Win32Apps\Output`).  
  
   Example:  
  
   ```  
   IntuneWinAppUtil.exe -c C:\Win32Apps\MyApp -s MyApp.msi -o C:\Win32Apps\Output  
   ```  
   
4. The Microsoft Win32 Content Prep Tool will create a .intunewin file in the specified output folder.  
   
## Deploying the Win32 App in Intune  
   
1. Sign in to the [Microsoft Endpoint Manager admin center](https://endpoint.microsoft.com/).  
2. Navigate to **Apps** > **All apps** > **Add**.  
3. In the *Select app type* pane, choose **Windows app (Win32)** and click **Select**.  
4. In the *App information* pane, click **Select app package file** and upload the .intunewin file you created earlier.  
5. Fill in the required app information, such as Name, Description, and Publisher.  
6. In the *Program* pane, provide the install and uninstall commands for your Win32 app.  
7. Configure the app's requirements, detection rules, and return codes in the respective panes.  
8. Click **Add** to finish creating the app.  
9. Assign the app to the desired user or device groups for deployment.  
   
That's it! You've successfully created a .intunewin file and deployed a Win32 app using Microsoft Intune.