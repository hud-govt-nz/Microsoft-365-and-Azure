<#
.SYNOPSIS
    Bulk Group Creation Script in Azure AD

.DESCRIPTION
    Import Groups from CSV and bulk create in Azure AD

.OUTPUTS
    Transcript file will be located in "C:Support\Log\GroupCreation_Script.log"

.NOTES
    Version:        1.0
    Author:         Ashley Forde
    Creation Date:  08.09.22

  
.EXAMPLE
  
#>

#Package Details
$PackageName = "GroupCreation_Script"
$Version = "1.0"
$OrgFolder = "HUD Tools"


#Root Directory & Logging
Start-Transcript -Path "C:Support\Log\$PackageName-Install.log" -Force -Append -Confirm:$False

##########################################################################################################################

##INSERT CODE HERE##

function Get-File {
    [cmdletBinding()]
    param(
        [Parameter()]
        [ValidateScript({Test-Path $_})]
        [String]
        $InitialDirectory
    )

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog 
    if($InitialDirectory){
    $FileBrowser.InitialDirectory = $InitialDirectory
    }
    else{
    $fileBrowser.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }   
    $FileBrowser.Filter = 'CSV (*.csv)|*.csv|All Files (*.*)|*.*'

    [void]$FileBrowser.ShowDialog()
    $FileBrowser.FileName
}

#Import File
$CSVFile = Get-File
$CSV = Import-CSV -path $CSVFile


#Define group type variables
$grouptype1 = "Standard"

#Get list of existing groups
$standardgroups = Get-AzureADGroup -All:$true

foreach ($group in $csv) {
    #Reset variable
    $groupexists = $false
    #Check group does not exist
    if (($standardgroups.DisplayName -contains $group.GroupName) -or ($dynamicgroups.DisplayName -contains $group.GroupName)) {
            write-host $group.GroupName "already exists"
            $groupexists = $true`
            
    if (($groupexists -eq $true) -and ($group.update -eq $true) -and ($group.grouptype -eq $grouptype2)) {
    
    $objectid- (Get-AzureADMSGroup -searchstring $group.GroupName).id 

    write-host $group.GroupName' is set to update. Updating group settings'

    Set-AzureADMSGroup `
    -id $objectid `
    -Description $group.description `
    -DisplayName $group.GroupName `
    -MailEnabled $false `
    -SecurityEnabled $true `
    -MailNickname $group.GroupName `
    -GroupTypes "DynamicMembership" `
    -MembershipRule $group.config #`
    #-MembershipRuleProcessingState "Paused"

        }

            elseif (($groupexists -eq $true) -and ($group.update -eq $true) -and ($group.grouptype -eq $grouptype1)) {

                $objectid- (Get-AzureADGroup -searchstring $group.GroupName).objectid 

                write-host $group.GroupName' is set to update. Updating group settings'

                Set-AzureADGroup `
                    -Description $group.description `
                    -DisplayName $group.GroupName `
                    -MailEnabled $false `
                    -SecurityEnabled $true `
                    -MailNickname "NotSet"
            }

    }

    elseif (($groupexists -eq $false) -and ($group.grouptype -eq $grouptype2)) {

        New-AzureADMSGroup `
            -Description $group.description `
            -DisplayName $group.GroupName `
            -MailEnabled $false `
            -SecurityEnabled $true `
            -MailNickname $group.group `
            -GroupTypes "DynamicMembership" `
            -MembershipRule $group.config #`
            #-MembershipRuleProcessingState "Paused"
    }

    elseif (($groupexists -eq $false) -and ($group.grouptype -eq $grouptype1)) {

        New-AzureADGroup `
            -Description $group.description `
            -DisplayName $group.GroupName `
            -MailEnabled $false `
            -SecurityEnabled $true `
            -MailNickname "NotSet"
            #-MembershipRuleProcessingState "Paused"
    }

    else { 
        write-host "Unknown issue:"$Error[0]
    }
}



##########################################################################################################################

Stop-Transcript
