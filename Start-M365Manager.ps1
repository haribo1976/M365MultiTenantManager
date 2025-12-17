<#
.SYNOPSIS
    M365 Multi-Tenant Manager - Entry Point
.DESCRIPTION
    PowerShell + WPF application for managing multiple Microsoft 365 tenants.
    Based on the IntuneManagement architecture pattern.
.PARAMETER Silent
    Run in silent/batch mode without UI
.PARAMETER SilentBatchFile
    Path to batch job definition JSON file
.PARAMETER TenantId
    Target tenant ID for silent operations
.PARAMETER AppId
    Application ID for authentication
.PARAMETER Secret
    Client secret for app authentication
.PARAMETER CertThumbprint
    Certificate thumbprint for app authentication
.PARAMETER ExportPath
    Output path for export operations
.EXAMPLE
    .\Start-M365Manager.ps1
    Launches the WPF UI
.EXAMPLE
    .\Start-M365Manager.ps1 -Silent -SilentBatchFile ".\batch.json" -TenantId "xxx" -AppId "yyy" -Secret "zzz"
    Runs batch operations without UI
.NOTES
    Version: 1.0.0
    Author: Fintel Services Division
#>

[CmdletBinding()]
param(
    [switch]$Silent,
    [string]$SilentBatchFile,
    [string]$TenantId,
    [string]$AppId,
    [string]$Secret,
    [string]$CertThumbprint,
    [string]$ExportPath
)

#region Initialization
$ErrorActionPreference = "Stop"
$script:ScriptRoot = $PSScriptRoot
$script:Version = "1.0.0"

# Unblock files if needed (security warning mitigation)
Get-ChildItem -Path $script:ScriptRoot -Recurse | Unblock-File -ErrorAction SilentlyContinue

# Set window title
$Host.UI.RawUI.WindowTitle = "M365 Multi-Tenant Manager v$($script:Version)"

# Create required directories
@("Export", "Logs", "Config") | ForEach-Object {
    $path = Join-Path $script:ScriptRoot $_
    if (-not (Test-Path $path)) {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
    }
}
#endregion

#region Module Loading
Write-Host "Loading M365 Multi-Tenant Manager v$($script:Version)..." -ForegroundColor Cyan

# Pre-load MSAL.PS globally
try {
    Import-Module MSAL.PS -Force -Global -ErrorAction Stop
    Write-Host "  MSAL.PS pre-loaded" -ForegroundColor Green
}
catch {
    Write-Host "  WARNING: MSAL.PS not available" -ForegroundColor Yellow
}

# Load core modules in order
$coreModules = @(
    "Core\Logging.psm1",
    "Core\Settings.psm1",
    "Core\Authentication.psm1",
    "Core\GraphAPI.psm1",
    "Core\Core.psm1"
)

foreach ($module in $coreModules) {
    $modulePath = Join-Path $script:ScriptRoot $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -Global
            Write-Host "  Loaded: $module" -ForegroundColor Green
        }
        catch {
            Write-Host "  Failed to load: $module - $_" -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "  Missing: $module" -ForegroundColor Yellow
    }
}

# Load extension modules
$extensionPath = Join-Path $script:ScriptRoot "Extensions"
if (Test-Path $extensionPath) {
    Get-ChildItem -Path $extensionPath -Directory | ForEach-Object {
        $extModule = Join-Path $_.FullName "$($_.Name).psm1"
        if (Test-Path $extModule) {
            try {
                Import-Module $extModule -Force -Global
                Write-Host "  Loaded Extension: $($_.Name)" -ForegroundColor Green
            }
            catch {
                Write-Host "  Failed to load extension: $($_.Name) - $_" -ForegroundColor Yellow
            }
        }
    }
}
#endregion

#region Main Execution
if ($Silent) {
    # Silent/Batch mode
    Write-Host "`nRunning in Silent Mode..." -ForegroundColor Cyan
    
    if (-not $SilentBatchFile) {
        Write-Host "Error: -SilentBatchFile required for silent mode" -ForegroundColor Red
        exit 1
    }
    
    if (-not (Test-Path $SilentBatchFile)) {
        Write-Host "Error: Batch file not found: $SilentBatchFile" -ForegroundColor Red
        exit 1
    }
    
    # Authenticate if credentials provided
    if ($TenantId) {
        $authParams = @{ TenantId = $TenantId }
        if ($AppId) { $authParams.AppId = $AppId }
        if ($Secret) { $authParams.Secret = $Secret }
        if ($CertThumbprint) { $authParams.CertThumbprint = $CertThumbprint }
        
        try {
            Connect-M365Tenant @authParams
            Write-Host "Authenticated to tenant: $TenantId" -ForegroundColor Green
        }
        catch {
            Write-Host "Authentication failed: $_" -ForegroundColor Red
            exit 1
        }
    }
    
    # Execute batch jobs
    try {
        $batchConfig = Get-Content $SilentBatchFile -Raw | ConvertFrom-Json
        Invoke-M365BatchJobs -Jobs $batchConfig.jobs -ExportPath $ExportPath
    }
    catch {
        Write-Host "Batch execution failed: $_" -ForegroundColor Red
        exit 1
    }
}
else {
    # UI Mode
    Write-Host "`nStarting UI..." -ForegroundColor Cyan
    
    # Load WPF assemblies
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms
    
    # Load and show main window
    try {
        $mainWindowPath = Join-Path $script:ScriptRoot "Core\UI\MainWindow.xaml"
        $mainWindowLogicPath = Join-Path $script:ScriptRoot "Core\UI\MainWindow.xaml.ps1"
        
        if (Test-Path $mainWindowLogicPath) {
            . $mainWindowLogicPath
            Show-MainWindow
        }
        else {
            Write-Host "UI files not found. Running in console mode." -ForegroundColor Yellow
            Write-Host "Use -Silent parameter for batch operations." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Failed to start UI: $_" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        exit 1
    }
}
#endregion
