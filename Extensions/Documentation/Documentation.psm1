<#
.SYNOPSIS
    Documentation Extension
.DESCRIPTION
    Functions for exporting and documenting tenant configurations.
#>

function Export-M365TenantData {
    <#
    .SYNOPSIS
        Export all data from a tenant
    .PARAMETER TenantId
        Tenant ID (uses current if not specified)
    .PARAMETER OutputPath
        Base output directory
    .PARAMETER Scope
        What to export (Users, Groups, Licenses, ConditionalAccess, All)
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId,
        
        [string]$OutputPath,
        
        [ValidateSet("Users", "Groups", "Licenses", "ConditionalAccess", "SecureScore", "All")]
        [string[]]$Scope = @("All")
    )
    
    if ($TenantId) {
        Switch-M365Tenant -TenantId $TenantId
    }
    
    $context = Get-M365AuthContext
    
    # Setup output directory
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Join-Path (Get-M365Setting -Name "ExportPath") (Get-Date -Format "yyyyMMdd_HHmmss")
    }
    
    $tenantFolder = Join-Path $OutputPath $context.TenantId
    
    if (-not (Test-Path $tenantFolder)) {
        New-Item -Path $tenantFolder -ItemType Directory -Force | Out-Null
    }
    
    Write-M365Log -Message "Exporting tenant data to: $tenantFolder" -Level Information -Component "Documentation"
    
    $exportResults = @{
        TenantId    = $context.TenantId
        ExportDate  = (Get-Date).ToString("o")
        OutputPath  = $tenantFolder
        Items       = @()
    }
    
    # Export based on scope
    $exportAll = $Scope -contains "All"
    
    # Users
    if ($exportAll -or $Scope -contains "Users") {
        try {
            $usersPath = Join-Path $tenantFolder "Users.json"
            Export-M365Users -Path $usersPath -Format JSON
            $exportResults.Items += @{ Type = "Users"; Path = $usersPath; Status = "Success" }
        }
        catch {
            $exportResults.Items += @{ Type = "Users"; Path = $null; Status = "Failed"; Error = $_.Exception.Message }
        }
    }
    
    # Groups
    if ($exportAll -or $Scope -contains "Groups") {
        try {
            $groupsPath = Join-Path $tenantFolder "Groups.json"
            Export-M365Groups -Path $groupsPath -Format JSON -IncludeMembers
            $exportResults.Items += @{ Type = "Groups"; Path = $groupsPath; Status = "Success" }
        }
        catch {
            $exportResults.Items += @{ Type = "Groups"; Path = $null; Status = "Failed"; Error = $_.Exception.Message }
        }
    }
    
    # Licenses
    if ($exportAll -or $Scope -contains "Licenses") {
        try {
            $licensesPath = Join-Path $tenantFolder "Licenses.json"
            Export-M365LicenseReport -Path $licensesPath -Format JSON
            $exportResults.Items += @{ Type = "Licenses"; Path = $licensesPath; Status = "Success" }
        }
        catch {
            $exportResults.Items += @{ Type = "Licenses"; Path = $null; Status = "Failed"; Error = $_.Exception.Message }
        }
    }
    
    # Conditional Access
    if ($exportAll -or $Scope -contains "ConditionalAccess") {
        try {
            $caPath = Join-Path $tenantFolder "ConditionalAccess"
            Export-M365ConditionalAccessPolicy -PolicyId "All" -Path $caPath
            $exportResults.Items += @{ Type = "ConditionalAccess"; Path = $caPath; Status = "Success" }
        }
        catch {
            $exportResults.Items += @{ Type = "ConditionalAccess"; Path = $null; Status = "Failed"; Error = $_.Exception.Message }
        }
    }
    
    # Secure Score
    if ($exportAll -or $Scope -contains "SecureScore") {
        try {
            $scorePath = Join-Path $tenantFolder "SecureScore.json"
            $score = Get-M365SecureScore
            $score | ConvertTo-Json -Depth 10 | Set-Content $scorePath -Encoding UTF8
            $exportResults.Items += @{ Type = "SecureScore"; Path = $scorePath; Status = "Success" }
        }
        catch {
            $exportResults.Items += @{ Type = "SecureScore"; Path = $null; Status = "Failed"; Error = $_.Exception.Message }
        }
    }
    
    # Save export manifest
    $manifestPath = Join-Path $tenantFolder "ExportManifest.json"
    $exportResults | ConvertTo-Json -Depth 10 | Set-Content $manifestPath -Encoding UTF8
    
    Write-M365Log -Message "Export completed. Manifest: $manifestPath" -Level Information -Component "Documentation"
    
    return [PSCustomObject]$exportResults
}

function Export-M365AllTenants {
    <#
    .SYNOPSIS
        Export data from all registered tenants
    .PARAMETER OutputPath
        Base output directory
    .PARAMETER Scope
        What to export
    #>
    [CmdletBinding()]
    param(
        [string]$OutputPath,
        
        [ValidateSet("Users", "Groups", "Licenses", "ConditionalAccess", "SecureScore", "All")]
        [string[]]$Scope = @("All")
    )
    
    $tenants = Get-M365RegisteredTenants
    
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Join-Path (Get-M365Setting -Name "ExportPath") "AllTenants_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }
    
    $results = @()
    
    foreach ($tenant in $tenants) {
        Write-M365Log -Message "Exporting: $($tenant.friendlyName)" -Level Information -Component "Documentation"
        
        try {
            $result = Export-M365TenantData -TenantId $tenant.id -OutputPath $OutputPath -Scope $Scope
            $results += $result
        }
        catch {
            Write-M365Log -Message "Failed to export $($tenant.friendlyName): $_" -Level Error -Component "Documentation"
            $results += @{
                TenantId = $tenant.id
                Status   = "Failed"
                Error    = $_.Exception.Message
            }
        }
    }
    
    # Create summary
    $summaryPath = Join-Path $OutputPath "ExportSummary.json"
    @{
        ExportDate   = (Get-Date).ToString("o")
        TenantCount  = $tenants.Count
        OutputPath   = $OutputPath
        Results      = $results
    } | ConvertTo-Json -Depth 10 | Set-Content $summaryPath -Encoding UTF8
    
    Write-M365Log -Message "All exports completed. Summary: $summaryPath" -Level Information -Component "Documentation"
    
    return $results
}

function New-M365MigrationTable {
    <#
    .SYNOPSIS
        Create a migration table for cross-tenant import
    .PARAMETER SourceTenantId
        Source tenant ID
    .PARAMETER TargetTenantId
        Target tenant ID
    .PARAMETER OutputPath
        Output file path
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceTenantId,
        
        [Parameter(Mandatory)]
        [string]$TargetTenantId,
        
        [string]$OutputPath
    )
    
    Write-M365Log -Message "Creating migration table" -Level Information -Component "Documentation"
    
    $migrationTable = @{
        SourceTenantId = $SourceTenantId
        TargetTenantId = $TargetTenantId
        CreatedDate    = (Get-Date).ToString("o")
        Groups         = @()
        Users          = @()
    }
    
    # Get source groups
    Switch-M365Tenant -TenantId $SourceTenantId
    $sourceGroups = Get-M365Groups -All
    
    # Get target groups
    Switch-M365Tenant -TenantId $TargetTenantId
    $targetGroups = Get-M365Groups -All
    
    foreach ($sourceGroup in $sourceGroups) {
        $targetMatch = $targetGroups | Where-Object { $_.displayName -eq $sourceGroup.displayName }
        
        $migrationTable.Groups += @{
            SourceId        = $sourceGroup.id
            SourceName      = $sourceGroup.displayName
            TargetId        = $targetMatch.id
            TargetName      = $targetMatch.displayName
            Mapped          = ($null -ne $targetMatch)
            NeedsCreation   = ($null -eq $targetMatch)
        }
    }
    
    if ([string]::IsNullOrEmpty($OutputPath)) {
        $OutputPath = Join-Path (Get-M365Setting -Name "ExportPath") "MigrationTable_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    }
    
    $migrationTable | ConvertTo-Json -Depth 10 | Set-Content $OutputPath -Encoding UTF8
    
    Write-M365Log -Message "Migration table created: $OutputPath" -Level Information -Component "Documentation"
    
    return $OutputPath
}

Export-ModuleMember -Function @(
    'Export-M365TenantData',
    'Export-M365AllTenants',
    'New-M365MigrationTable'
)
