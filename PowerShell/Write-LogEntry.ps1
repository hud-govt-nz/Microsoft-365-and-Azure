function Write-LogEntry {
    <#
    .SYNOPSIS
    Writes a log entry to a log file with the specified value and severity.

    .DESCRIPTION
    This function writes a log entry to a log file with the specified value and severity. 
    It constructs the log entry with timestamp, date, context, and other relevant information.

    .PARAMETER Value
    [string] Value added to the log file.

    .PARAMETER Severity
    [int] Severity for the log entry. 1 for Informational, 2 for Warning, and 3 for Error.

    .PARAMETER LogFilePath
    [string] Path of the log file to write the entry to.

    .EXAMPLE
    Write-LogEntry -Value "This is an informational message" -Severity 1 -LogFilePath "C:\Logs\custom.log"

    Writes an informational log entry with the specified value to the custom log file path.
    #>

    param(
        [Parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [string]$value,
        [Parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateSet("1", "2", "3")]
        [int]$severity,
        [Parameter(Mandatory = $true, HelpMessage = "Path of the log file that the entry will be written to.")]
        [string]$logFilePath
    )

    # Get the current time, date, and context
    $time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $date = (Get-Date -Format "MM-dd-yyyy")
    $context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    
    # Construct the log entry text using string interpolation
    $logText = "<![LOG[$($value)]LOG]!><time=""$($time)"" date=""$($date)"" component=""$($logFilePath)"" context=""$($context)"" type=""$($severity)"" thread=""$($PID)"" file="""">"

    # Write the log entry to the log file, handling errors with -ErrorAction Stop
    Out-File -InputObject $logText -Append -NoClobber -Encoding Default -FilePath $logFilePath -ErrorAction Stop

    # Output the log message based on severity
    if ($severity -eq 1) {
        Write-Output -Message $value
    }
    elseif ($severity -eq 3) {
        Write-Output -Message $value
    }
}