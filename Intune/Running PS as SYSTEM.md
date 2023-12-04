Title: Setting up PsTools and Running PowerShell as NT\SYSTEM on Your Local Machine

## Introduction
PsTools is a suite of command-line utilities that allow you to administer your local and remote systems. One of the tools included in this suite is PsExec, which enables you to execute processes on remote systems and redirect console applications' input/output to the local system. In this tutorial, we'll guide you through setting up PsTools on your local machine and using it to run PowerShell as `NT\SYSTEM` which has high-level privileges on the machine.

## Pre-requisites
- A Windows operating system
- Administrator privileges on your local machine

## Setting up PsTools

1. **Download PsTools Suite**:
   - Navigate to the [PsTools webpage](https://docs.microsoft.com/en-us/sysinternals/downloads/pstools) on Microsoft’s Sysinternals site.
   - Click on the `Download PsTools` link to download the zip file.

2. **Extract PsTools Suite**:
   - Locate the downloaded zip file, usually in your Downloads folder.
   - Right-click the zip file and select `Extract All` to extract the files to a directory of your choice.

3. **Add PsTools to System Path**:
   - Right-click `This PC` or `My Computer` on your desktop or in File Explorer, and then click `Properties`.
   - Click on `Advanced system settings`.
   - Click on `Environment Variables`.
   - Under System Variables, scroll down and select `Path`, then click on `Edit`.
   - Click on `New` and add the path to the directory where you extracted PsTools.
   - Click `OK` to close each window.

## Running PowerShell as NT\SYSTEM

1. **Open Command Prompt as Administrator**:
   - Search for `Command Prompt` in the start menu, right-click on it, and select `Run as administrator`.

2. **Run PsExec**:
   - In the command prompt, type the following command and press `Enter`:
     ```shell
     psexec -i -s powershell.exe
     ```
   - The `-i` flag allows interaction with the new process, and the `-s` flag runs the new process as the `NT\SYSTEM` user.

3. **Verify the User**:
   - In the new PowerShell window, type the following command and press `Enter` to verify you are running as `NT\SYSTEM`:
     ```shell
     whoami
     ```

Now, you are running PowerShell with `NT\SYSTEM` privileges, and you can execute commands with high-level permissions on your local machine. Remember to exercise caution when running commands as `NT\SYSTEM` as you have the highest level of access on the machine, and it's easy to cause damage if you're not sure about what you are doing.
