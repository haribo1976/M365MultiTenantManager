<#
.SYNOPSIS
    Core Module
.DESCRIPTION
    Main module that provides core functionality and batch processing.
    This module depends on Logging, Settings, Authentication, and GraphAPI modules.
#>

#region Module Variables
$script:RegisteredTenants = @()
$script:TenantsFilePath = $null
#endregion

#region Tenant Registry Functions
function Initialize-M365TenantRegistry {
    <#
    .SYNOPSIS
        Load registered tenants from file
    #>
    $script:TenantsFilePath = Join-Path $PSScriptRoot "..\Config\Tenants.json"
    
    if (Test-Path $script:TenantsFilePath) {
        try {
            $content = Get-Content $script:TenantsFilePath -Raw | ConvertFrom-Json
            $script:RegisteredTenants = @($content.tenants)
            Write-M365Log -Message "Loaded $($script:RegisteredTenants.Count) registered tenants" -Level Debug -Component "Core"
        }
        catch {
            Write-M365Log -Message "Failed to load tenants file: $_" -Level Warning -Component "Core"
            $script:RegisteredTenants = @()
        }
    }
    else {
        $script:RegisteredTenants = @()
    }
}

function Get-M365RegisteredTenants {
    <#
    .SYNOPSIS
        Get all registered tenants
    #>
    if ($script:RegisteredTenants.Count -eq 0) {
        Initialize-M365TenantRegistry
    }
    
    return $script:RegisteredTenants
}

function Register-M365Tenant {
    <#
    .SYNOPSIS
        Register a new tenant for management
    .PARAMETER TenantId
        Tenant ID (GUID)
    .PARAMETER FriendlyName
        Display name for the tenant
    .PARAMETER Tags
        Tags for grouping/filtering
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,
        
        [string]$FriendlyName,
        
        [string[]]$Tags = @()
    )
    
    # Check if already registered
    $existing = $script:RegisteredTenants | Where-Object { $_.id -eq $TenantId }
    if ($existing) {
        Write-M365Log -Message "Tenant already registered: $TenantId" -Level Warning -Component "Core"
        return $existing
    }
    
    # Connect to get tenant info
    Write-M365Log -Message "Connecting to tenant for registration: $TenantId" -Level Information -Component "Core"
    
    try {
        Connect-M365Tenant -TenantId $TenantId
        
        # Get organization info
        $org = Invoke-M365GraphRequest -Method GET -Uri "/organization"
        $orgInfo = $org.value[0]
        
        # Get domains
        $domains = Invoke-M365GraphRequest -Method GET -Uri "/domains"
        $primaryDomain = ($domains.value | Where-Object { $_.isDefault }).id
        
        $tenant = [PSCustomObject]@{
            id             = $TenantId
            displayName    = $orgInfo.displayName
            friendlyName   = if ($FriendlyName) { $FriendlyName } else { $orgInfo.displayName }
            primaryDomain  = $primaryDomain
            tags           = $Tags
            registeredAt   = (Get-Date).ToString("o")
            lastAccessedAt = (Get-Date).ToString("o")
        }
        
        $script:RegisteredTenants += $tenant
        Save-M365TenantRegistry
        
        Write-M365Log -Message "Tenant registered: $($tenant.displayName)" -Level Information -Component "Core"
        
        return $tenant
    }
    catch {
        Write-M365Log -Message "Failed to register tenant: $_" -Level Error -Component "Core"
        throw
    }
}

function Unregister-M365Tenant {
    <#
    .SYNOPSIS
        Remove a tenant from registry
    .PARAMETER TenantId
        Tenant ID to remove
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    
    $script:RegisteredTenants = @($script:RegisteredTenants | Where-Object { $_.id -ne $TenantId })
    Save-M365TenantRegistry
    
    # Also clear cached token
    Disconnect-M365Tenant
    
    Write-M365Log -Message "Tenant unregistered: $TenantId" -Level Information -Component "Core"
}

function Save-M365TenantRegistry {
    <#
    .SYNOPSIS
        Save registered tenants to file
    #>
    try {
        $configDir = Split-Path $script:TenantsFilePath -Parent
        if (-not (Test-Path $configDir)) {
            New-Item -Path $configDir -ItemType Directory -Force | Out-Null
        }
        
        $data = @{
            version  = "1.0"
            updated  = (Get-Date).ToString("o")
            tenants  = $script:RegisteredTenants
        }
        
        $data | ConvertTo-Json -Depth 10 | Set-Content $script:TenantsFilePath -Encoding UTF8
        Write-M365Log -Message "Tenant registry saved" -Level Debug -Component "Core"
    }
    catch {
        Write-M365Log -Message "Failed to save tenant registry: $_" -Level Error -Component "Core"
    }
}

function Update-M365TenantAccess {
    <#
    .SYNOPSIS
        Update last accessed timestamp for a tenant
    .PARAMETER TenantId
        Tenant ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    
    $tenant = $script:RegisteredTenants | Where-Object { $_.id -eq $TenantId }
    if ($tenant) {
        $tenant.lastAccessedAt = (Get-Date).ToString("o")
        Save-M365TenantRegistry
    }
}
#endregion

#region Batch Processing Functions
function Invoke-M365BatchJobs {
    <#
    .SYNOPSIS
        Execute batch jobs from configuration
    .PARAMETER Jobs
        Array of job definitions
    .PARAMETER ExportPath
        Override export path
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Jobs,
        
        [string]$ExportPath
    )
    
    Write-M365Log -Message "Starting batch execution: $($Jobs.Count) jobs" -Level Information -Component "Batch"
    
    $results = @()
    
    foreach ($job in $Jobs) {
        Write-M365Log -Message "Executing job: $($job.name)" -Level Information -Component "Batch"
        
        try {
            $jobResult = [PSCustomObject]@{
                JobName   = $job.name
                Action    = $job.action
                Status    = "Running"
                StartTime = Get-Date
                EndTime   = $null
                Results   = @()
                Error     = $null
            }
            
            # Determine target tenants
            $targetTenants = @()
            if ($job.tenants -contains "*") {
                $targetTenants = Get-M365RegisteredTenants
            }
            else {
                $targetTenants = Get-M365RegisteredTenants | Where-Object { $_.id -in $job.tenants }
            }
            
            foreach ($tenant in $targetTenants) {
                Write-M365Log -Message "  Processing tenant: $($tenant.displayName)" -Level Information -Component "Batch"
                
                try {
                    Switch-M365Tenant -TenantId $tenant.id
                    
                    switch ($job.action.ToLower()) {
                        "export" {
                            $outputPath = $job.outputPath
                            if ($ExportPath) { $outputPath = $ExportPath }
                            $outputPath = $outputPath -replace "{TenantName}", $tenant.friendlyName
                            $outputPath = $outputPath -replace "{DateTime}", (Get-Date -Format "yyyyMMdd_HHmmss")
                            
                            # Call export functions based on scope
                            # This would call extension module functions
                            $jobResult.Results += @{
                                TenantId = $tenant.id
                                Status   = "Exported"
                                Path     = $outputPath
                            }
                        }
                        "compare" {
                            # Compare operation
                            $jobResult.Results += @{
                                SourceTenant = $job.sourceTenant
                                TargetTenant = $job.targetTenant
                                Status       = "Compared"
                            }
                        }
                        default {
                            Write-M365Log -Message "Unknown action: $($job.action)" -Level Warning -Component "Batch"
                        }
                    }
                }
                catch {
                    Write-M365Log -Message "Error processing tenant $($tenant.id): $_" -Level Error -Component "Batch"
                    $jobResult.Results += @{
                        TenantId = $tenant.id
                        Status   = "Error"
                        Error    = $_.Exception.Message
                    }
                }
            }
            
            $jobResult.Status = "Completed"
            $jobResult.EndTime = Get-Date
        }
        catch {
            $jobResult.Status = "Failed"
            $jobResult.Error = $_.Exception.Message
            $jobResult.EndTime = Get-Date
            Write-M365Log -Message "Job failed: $($job.name) - $_" -Level Error -Component "Batch"
        }
        
        $results += $jobResult
    }
    
    Write-M365Log -Message "Batch execution completed" -Level Information -Component "Batch"
    
    return $results
}
#endregion

#region Utility Functions
function Get-M365ManagerVersion {
    <#
    .SYNOPSIS
        Get application version
    #>
    return "1.0.0"
}

function Test-M365Prerequisites {
    <#
    .SYNOPSIS
        Check if all prerequisites are met
    #>
    $results = @{
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        PowerShell7       = $PSVersionTable.PSVersion.Major -ge 7
        MSALAvailable     = $false
        WPFAvailable      = $false
    }
    
    # Check MSAL
    if (Get-Module -ListAvailable -Name MSAL.PS) {
        $results.MSALAvailable = $true
    }
    
    # Check WPF
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
        $results.WPFAvailable = $true
    }
    catch { }
    
    return [PSCustomObject]$results
}

function Show-M365ManagerHelp {
    <#
    .SYNOPSIS
        Display help information
    #>
    $help = @"
M365 Multi-Tenant Manager v$(Get-M365ManagerVersion)
========================================

Usage:
  .\Start-M365Manager.ps1                    # Launch UI
  .\Start-M365Manager.ps1 -Silent ...        # Batch mode

Parameters:
  -Silent              Run without UI
  -SilentBatchFile     Path to batch job JSON
  -TenantId            Target tenant ID
  -AppId               Application ID
  -Secret              Client secret
  -CertThumbprint      Certificate thumbprint
  -ExportPath          Output path override

Commands (in console):
  Connect-M365Tenant      Connect to a tenant
  Get-M365RegisteredTenants  List registered tenants
  Register-M365Tenant     Add a new tenant
  Switch-M365Tenant       Switch tenant context

For more information, see the documentation.
"@
    
    Write-Host $help
}
#endregion

#region Module Initialization
Initialize-M365TenantRegistry
#endregion

Export-ModuleMember -Function @(
    'Get-M365RegisteredTenants',
    'Register-M365Tenant',
    'Unregister-M365Tenant',
    'Update-M365TenantAccess',
    'Invoke-M365BatchJobs',
    'Get-M365ManagerVersion',
    'Test-M365Prerequisites',
    'Show-M365ManagerHelp'
)
