<#
.SYNOPSIS
    Settings Module
.DESCRIPTION
    Manages application settings stored in JSON format.
    Provides functions to read, write, and reset settings.
#>

#region Module Variables
$script:Settings = $null
$script:SettingsPath = $null
$script:DefaultSettings = @{
    # Authentication
    DefaultAppId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph PowerShell
    CustomAppId = $null
    UseDefaultPermissions = $true
    
    # API Settings
    GraphVersion = "v1.0"
    GraphBetaEndpoints = @("secureScores", "conditionalAccess")
    RequestTimeout = 30
    MaxRetries = 3
    
    # Paths
    ExportPath = ".\Export"
    LogPath = ".\Logs"
    
    # Logging
    LogLevel = "Information"
    
    # UI Settings
    Theme = "Light"
    AutoRefreshInterval = 300
    ShowGuestUsers = $true
    ShowDisabledUsers = $false
    
    # Cache
    CacheExpiryMinutes = 15
    EnableCaching = $true
    
    # Display
    PageSize = 100
    DateFormat = "yyyy-MM-dd HH:mm:ss"
}
#endregion

#region Public Functions
function Initialize-M365Settings {
    <#
    .SYNOPSIS
        Initialize settings from file or defaults
    .PARAMETER SettingsPath
        Path to settings JSON file
    #>
    [CmdletBinding()]
    param(
        [string]$SettingsPath
    )
    
    if ([string]::IsNullOrEmpty($SettingsPath)) {
        $SettingsPath = Join-Path $PSScriptRoot "..\Config\Settings.json"
    }
    
    $script:SettingsPath = $SettingsPath
    
    if (Test-Path $SettingsPath) {
        try {
            $fileContent = Get-Content $SettingsPath -Raw | ConvertFrom-Json
            $script:Settings = @{}
            
            # Merge with defaults (file settings override defaults)
            foreach ($key in $script:DefaultSettings.Keys) {
                if ($null -ne $fileContent.$key) {
                    $script:Settings[$key] = $fileContent.$key
                }
                else {
                    $script:Settings[$key] = $script:DefaultSettings[$key]
                }
            }
            
            Write-M365Log -Message "Settings loaded from: $SettingsPath" -Level Debug
        }
        catch {
            Write-M365Log -Message "Failed to load settings, using defaults: $_" -Level Warning
            $script:Settings = $script:DefaultSettings.Clone()
        }
    }
    else {
        Write-M365Log -Message "Settings file not found, using defaults" -Level Debug
        $script:Settings = $script:DefaultSettings.Clone()
        Save-M365Settings
    }
}

function Get-M365Setting {
    <#
    .SYNOPSIS
        Get a setting value
    .PARAMETER Name
        Setting name
    .PARAMETER Default
        Default value if setting not found
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        
        [object]$Default = $null
    )
    
    if (-not $script:Settings) {
        Initialize-M365Settings
    }
    
    if ($script:Settings.ContainsKey($Name)) {
        return $script:Settings[$Name]
    }
    
    return $Default
}

function Set-M365Setting {
    <#
    .SYNOPSIS
        Set a setting value
    .PARAMETER Name
        Setting name
    .PARAMETER Value
        Setting value
    .PARAMETER NoSave
        Don't persist to file immediately
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        
        [Parameter(Mandatory)]
        [object]$Value,
        
        [switch]$NoSave
    )
    
    if (-not $script:Settings) {
        Initialize-M365Settings
    }
    
    $script:Settings[$Name] = $Value
    Write-M365Log -Message "Setting updated: $Name" -Level Debug
    
    if (-not $NoSave) {
        Save-M365Settings
    }
}

function Get-M365AllSettings {
    <#
    .SYNOPSIS
        Get all current settings
    #>
    if (-not $script:Settings) {
        Initialize-M365Settings
    }
    
    return $script:Settings.Clone()
}

function Save-M365Settings {
    <#
    .SYNOPSIS
        Save current settings to file
    #>
    if (-not $script:Settings) {
        return
    }
    
    try {
        $configDir = Split-Path $script:SettingsPath -Parent
        if (-not (Test-Path $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        }
        
        $script:Settings | ConvertTo-Json -Depth 10 | Set-Content $script:SettingsPath -Encoding UTF8
        Write-M365Log -Message "Settings saved to: $($script:SettingsPath)" -Level Debug
    }
    catch {
        Write-M365Log -Message "Failed to save settings: $_" -Level Error
    }
}

function Reset-M365Settings {
    <#
    .SYNOPSIS
        Reset all settings to defaults
    #>
    $script:Settings = $script:DefaultSettings.Clone()
    Save-M365Settings
    Write-M365Log -Message "Settings reset to defaults" -Level Information
}

function Get-M365SettingsPath {
    <#
    .SYNOPSIS
        Get path to settings file
    #>
    return $script:SettingsPath
}
#endregion

#region Module Initialization
Initialize-M365Settings
#endregion

Export-ModuleMember -Function @(
    'Initialize-M365Settings',
    'Get-M365Setting',
    'Set-M365Setting',
    'Get-M365AllSettings',
    'Save-M365Settings',
    'Reset-M365Settings',
    'Get-M365SettingsPath'
)
