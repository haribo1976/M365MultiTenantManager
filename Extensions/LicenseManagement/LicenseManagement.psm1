<#
.SYNOPSIS
    License Management Extension
.DESCRIPTION
    Functions for viewing and managing licenses across tenants.
#>

function Get-M365Licenses {
    <#
    .SYNOPSIS
        Get all licenses in current tenant
    #>
    [CmdletBinding()]
    param()
    
    $skus = Invoke-M365GraphRequest -Method GET -Uri "/subscribedSkus"
    
    return $skus.value | ForEach-Object {
        [PSCustomObject]@{
            SkuId          = $_.skuId
            SkuPartNumber  = $_.skuPartNumber
            DisplayName    = Get-LicenseFriendlyName -SkuPartNumber $_.skuPartNumber
            Total          = $_.prepaidUnits.enabled
            Assigned       = $_.consumedUnits
            Available      = $_.prepaidUnits.enabled - $_.consumedUnits
            Warning        = $_.prepaidUnits.warning
            Suspended      = $_.prepaidUnits.suspended
            CapabilityStatus = $_.capabilityStatus
            AppliesTo      = $_.appliesTo
        }
    }
}

function Get-M365LicenseSummary {
    <#
    .SYNOPSIS
        Get license summary for one or all tenants
    .PARAMETER AllTenants
        Get summary across all registered tenants
    #>
    [CmdletBinding()]
    param(
        [switch]$AllTenants
    )
    
    $results = @()
    
    if ($AllTenants) {
        $tenants = Get-M365RegisteredTenants
        
        foreach ($tenant in $tenants) {
            Write-M365Log -Message "Getting licenses for: $($tenant.friendlyName)" -Level Debug -Component "LicenseMgmt"
            
            try {
                Switch-M365Tenant -TenantId $tenant.id
                $licenses = Get-M365Licenses
                
                foreach ($lic in $licenses) {
                    $results += [PSCustomObject]@{
                        TenantId      = $tenant.id
                        TenantName    = $tenant.friendlyName
                        SkuPartNumber = $lic.SkuPartNumber
                        DisplayName   = $lic.DisplayName
                        Total         = $lic.Total
                        Assigned      = $lic.Assigned
                        Available     = $lic.Available
                        Utilisation   = if ($lic.Total -gt 0) { [math]::Round(($lic.Assigned / $lic.Total) * 100, 1) } else { 0 }
                    }
                }
            }
            catch {
                Write-M365Log -Message "Failed to get licenses for $($tenant.friendlyName): $_" -Level Warning -Component "LicenseMgmt"
            }
        }
    }
    else {
        $licenses = Get-M365Licenses
        $context = Get-M365AuthContext
        
        foreach ($lic in $licenses) {
            $results += [PSCustomObject]@{
                TenantId      = $context.TenantId
                TenantName    = "Current"
                SkuPartNumber = $lic.SkuPartNumber
                DisplayName   = $lic.DisplayName
                Total         = $lic.Total
                Assigned      = $lic.Assigned
                Available     = $lic.Available
                Utilisation   = if ($lic.Total -gt 0) { [math]::Round(($lic.Assigned / $lic.Total) * 100, 1) } else { 0 }
            }
        }
    }
    
    return $results
}

function Compare-M365Licenses {
    <#
    .SYNOPSIS
        Compare licenses between two tenants
    .PARAMETER SourceTenantId
        Source tenant ID
    .PARAMETER TargetTenantId
        Target tenant ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceTenantId,
        
        [Parameter(Mandatory)]
        [string]$TargetTenantId
    )
    
    # Get source licenses
    Switch-M365Tenant -TenantId $SourceTenantId
    $sourceLicenses = Get-M365Licenses
    
    # Get target licenses
    Switch-M365Tenant -TenantId $TargetTenantId
    $targetLicenses = Get-M365Licenses
    
    $comparison = @()
    
    # All unique SKUs
    $allSkus = ($sourceLicenses.SkuPartNumber + $targetLicenses.SkuPartNumber) | Select-Object -Unique
    
    foreach ($sku in $allSkus) {
        $source = $sourceLicenses | Where-Object { $_.SkuPartNumber -eq $sku }
        $target = $targetLicenses | Where-Object { $_.SkuPartNumber -eq $sku }
        
        $status = "Match"
        if (-not $source) { $status = "TargetOnly" }
        elseif (-not $target) { $status = "SourceOnly" }
        elseif ($source.Total -ne $target.Total) { $status = "Different" }
        
        $comparison += [PSCustomObject]@{
            SkuPartNumber      = $sku
            DisplayName        = Get-LicenseFriendlyName -SkuPartNumber $sku
            Status             = $status
            SourceTotal        = $source.Total
            SourceAssigned     = $source.Assigned
            TargetTotal        = $target.Total
            TargetAssigned     = $target.Assigned
        }
    }
    
    return $comparison
}

function Export-M365LicenseReport {
    <#
    .SYNOPSIS
        Export license report to file
    .PARAMETER Path
        Output file path
    .PARAMETER AllTenants
        Include all registered tenants
    .PARAMETER Format
        Output format (CSV or JSON)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        
        [switch]$AllTenants,
        
        [ValidateSet("CSV", "JSON")]
        [string]$Format = "CSV"
    )
    
    Write-M365Log -Message "Exporting license report to: $Path" -Level Information -Component "LicenseMgmt"
    
    $licenses = Get-M365LicenseSummary -AllTenants:$AllTenants
    
    if ($Format -eq "CSV") {
        $licenses | Export-Csv -Path $Path -NoTypeInformation
    }
    else {
        $licenses | ConvertTo-Json -Depth 10 | Set-Content $Path -Encoding UTF8
    }
    
    Write-M365Log -Message "License report exported" -Level Information -Component "LicenseMgmt"
    
    return $Path
}

function Get-LicenseFriendlyName {
    <#
    .SYNOPSIS
        Convert SKU part number to friendly name
    #>
    param([string]$SkuPartNumber)
    
    # Common SKU mappings
    $friendlyNames = @{
        "ENTERPRISEPREMIUM"     = "Microsoft 365 E5"
        "ENTERPRISEPACK"        = "Microsoft 365 E3"
        "SPE_E3"               = "Microsoft 365 E3"
        "SPE_E5"               = "Microsoft 365 E5"
        "SPB"                  = "Microsoft 365 Business Premium"
        "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Premium"
        "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
        "EXCHANGESTANDARD"      = "Exchange Online Plan 1"
        "EXCHANGEENTERPRISE"    = "Exchange Online Plan 2"
        "EMS"                  = "Enterprise Mobility + Security E3"
        "EMSPREMIUM"           = "Enterprise Mobility + Security E5"
        "ATP_ENTERPRISE"       = "Microsoft Defender for Office 365 P1"
        "THREAT_INTELLIGENCE"  = "Microsoft Defender for Office 365 P2"
        "IDENTITY_THREAT_PROTECTION" = "Microsoft 365 E5 Security"
        "AAD_PREMIUM"          = "Azure AD Premium P1"
        "AAD_PREMIUM_P2"       = "Azure AD Premium P2"
        "INTUNE_A"             = "Microsoft Intune"
        "WIN10_PRO_ENT_SUB"    = "Windows 10/11 Enterprise E3"
        "VISIOCLIENT"          = "Visio Plan 2"
        "PROJECTPROFESSIONAL"  = "Project Plan 3"
        "POWER_BI_PRO"         = "Power BI Pro"
        "STREAM"               = "Microsoft Stream"
        "FLOW_FREE"            = "Power Automate Free"
        "POWERAPPS_VIRAL"      = "Power Apps Free"
        "TEAMS_EXPLORATORY"    = "Microsoft Teams Exploratory"
        "M365_F1"              = "Microsoft 365 F1"
        "SPE_F1"               = "Microsoft 365 F3"
    }
    
    if ($friendlyNames.ContainsKey($SkuPartNumber)) {
        return $friendlyNames[$SkuPartNumber]
    }
    
    # Return original if no mapping found
    return $SkuPartNumber
}

Export-ModuleMember -Function @(
    'Get-M365Licenses',
    'Get-M365LicenseSummary',
    'Compare-M365Licenses',
    'Export-M365LicenseReport',
    'Get-LicenseFriendlyName'
)
