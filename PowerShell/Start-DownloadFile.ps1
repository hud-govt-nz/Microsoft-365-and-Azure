function Start-DownloadFile {
    <#
    .SYNOPSIS
    Downloads a file from a specified URL to a specified path.
    
    .PARAMETER URL
    The URL of the file to be downloaded.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$URL,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    begin {
        [System.Net.WebClient]$WebClient = New-Object -TypeName System.Net.WebClient
    }
    process {
        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        $WebClient.DownloadFile($URL,(Join-Path -Path $Path -ChildPath $Name))
    }
    end {
        $WebClient.Dispose()
    }
}