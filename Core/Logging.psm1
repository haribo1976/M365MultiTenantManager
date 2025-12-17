<#
.SYNOPSIS
    Logging Module - CMTrace Compatible
.DESCRIPTION
    Provides logging functions that output in CMTrace-compatible format.
    Logs can be viewed with CMTrace, OneTrace, or any text editor.
#>

#region Module Variables
$script:LogPath = $null
$script:LogFile = $null
$script:LogLevel = "Information"
$script:LogLevels = @{
    "Debug" = 0
    "Information" = 1
    "Warning" = 2
    "Error" = 3
}
#endregion

#region Public Functions
function Initialize-M365Logging {
    <#
    .SYNOPSIS
        Initialize logging for the session
    .PARAMETER LogPath
        Directory for log files
    .PARAMETER LogLevel
        Minimum log level (Debug, Information, Warning, Error)
    #>
    [CmdletBinding()]
    param(
        [string]$LogPath,
        [ValidateSet("Debug", "Information", "Warning", "Error")]
        [string]$LogLevel = "Information"
    )
    
    if ([string]::IsNullOrEmpty($LogPath)) {
        $LogPath = Join-Path $PSScriptRoot "..\Logs"
    }
    
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    $script:LogPath = $LogPath
    $script:LogLevel = $LogLevel
    $script:LogFile = Join-Path $LogPath "M365Manager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    
    Write-M365Log -Message "Logging initialized. Level: $LogLevel" -Level Information
}

function Write-M365Log {
    <#
    .SYNOPSIS
        Write a log entry in CMTrace format
    .PARAMETER Message
        Log message
    .PARAMETER Level
        Log level (Debug, Information, Warning, Error)
    .PARAMETER Component
        Component name for filtering
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet("Debug", "Information", "Warning", "Error")]
        [string]$Level = "Information",
        
        [string]$Component = "M365Manager"
    )
    
    # Check if we should log this level
    if ($script:LogLevels[$Level] -lt $script:LogLevels[$script:LogLevel]) {
        return
    }
    
    # Initialize if not done
    if (-not $script:LogFile) {
        Initialize-M365Logging
    }
    
    # CMTrace log format
    $timeGenerated = Get-Date -Format "HH:mm:ss.fff"
    $dateGenerated = Get-Date -Format "MM-dd-yyyy"
    
    # CMTrace severity: 1=Info, 2=Warning, 3=Error
    $severity = switch ($Level) {
        "Debug" { 1 }
        "Information" { 1 }
        "Warning" { 2 }
        "Error" { 3 }
        default { 1 }
    }
    
    # Build CMTrace format string
    $logEntry = "<![LOG[$Message]LOG]!>" +
                "<time=`"$timeGenerated`" " +
                "date=`"$dateGenerated`" " +
                "component=`"$Component`" " +
                "context=`"`" " +
                "type=`"$severity`" " +
                "thread=`"$PID`" " +
                "file=`"`">"
    
    # Write to file
    try {
        Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write log: $_"
    }
    
    # Also write to console with color
    $color = switch ($Level) {
        "Debug" { "Gray" }
        "Information" { "White" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        default { "White" }
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-M365LogPath {
    <#
    .SYNOPSIS
        Get current log file path
    #>
    return $script:LogFile
}

function Open-M365Log {
    <#
    .SYNOPSIS
        Open current log file in default viewer
    #>
    if ($script:LogFile -and (Test-Path $script:LogFile)) {
        Start-Process $script:LogFile
    }
    else {
        Write-Warning "No log file available"
    }
}
#endregion

#region Module Initialization
# Auto-initialize with defaults when module loads
Initialize-M365Logging
#endregion

Export-ModuleMember -Function @(
    'Initialize-M365Logging',
    'Write-M365Log',
    'Get-M365LogPath',
    'Open-M365Log'
)
