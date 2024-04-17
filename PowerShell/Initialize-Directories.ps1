function Initialize-Directories {
    <#
    .SYNOPSIS
    Initializes the directory structure for the home folder.

    .PARAMETER HomeFolder
    The path of the home folder

    .EXAMPLE
    Initialize-Directories -HomeFolder "C:\MyFolder"
    Creates the directory structure for the specified home folder.

    .OUTPUTS
    PSCustomObject
    Returns a custom object containing the folder paths.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$HomeFolder
    )

    if (Test-Path -Path $HomeFolder) {
        # Force creating 00_Staging folder at a minimum if it is missing
        New-Item -Path $HomeFolder -Name "00_Staging" -ItemType "directory" -Force -Confirm:$false | Out-Null
    }
    else {
        New-Item -Path $HomeFolder -ItemType "directory" -Force -Confirm:$false
        foreach ($subFolder in "00_Staging","01_Logs","02_Validation") {
            New-Item -Path $HomeFolder -Name $subFolder -ItemType "directory" -Force -Confirm:$false
        }
    }

    $StagingFolder = Join-Path -Path $HomeFolder -ChildPath "00_Staging"
    $LogsFolder = Join-Path -Path $HomeFolder -ChildPath "01_Logs"
    $ValidationFolder = Join-Path -Path $HomeFolder -ChildPath "02_Validation"

    return [PSCustomObject]@{
        HomeFolder = $HomeFolder
        StagingFolder = $StagingFolder
        LogsFolder = $LogsFolder
        ValidationFolder = $ValidationFolder
    }
}