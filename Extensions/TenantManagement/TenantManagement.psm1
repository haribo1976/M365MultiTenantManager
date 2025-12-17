<#
.SYNOPSIS
    Tenant Management Extension
.DESCRIPTION
    Functions for managing tenant registration, health, and information.
#>

function Get-M365TenantInfo {
    <#
    .SYNOPSIS
        Get detailed information about a tenant
    .PARAMETER TenantId
        Tenant ID (uses current if not specified)
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId
    )
    
    if ($TenantId) {
        Switch-M365Tenant -TenantId $TenantId
    }
    
    $org = Invoke-M365GraphRequest -Method GET -Uri "/organization"
    $domains = Invoke-M365GraphRequest -Method GET -Uri "/domains"
    
    return [PSCustomObject]@{
        TenantId         = $org.value[0].id
        DisplayName      = $org.value[0].displayName
        CreatedDateTime  = $org.value[0].createdDateTime
        TenantType       = $org.value[0].tenantType
        VerifiedDomains  = $domains.value | Where-Object { $_.isVerified } | Select-Object -ExpandProperty id
        PrimaryDomain    = ($domains.value | Where-Object { $_.isDefault }).id
        TechnicalContact = $org.value[0].technicalNotificationMails
    }
}

function Get-M365TenantHealth {
    <#
    .SYNOPSIS
        Get health status for a tenant
    .PARAMETER TenantId
        Tenant ID (uses current if not specified)
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId
    )
    
    if ($TenantId) {
        Switch-M365Tenant -TenantId $TenantId
    }
    
    try {
        $health = Invoke-M365GraphRequest -Method GET -Uri "/admin/serviceAnnouncement/healthOverviews"
        
        $summary = @{
            Healthy   = 0
            Advisory  = 0
            Incident  = 0
            Services  = @()
        }
        
        foreach ($service in $health.value) {
            $summary.Services += [PSCustomObject]@{
                Service = $service.service
                Status  = $service.status
            }
            
            switch ($service.status) {
                "serviceOperational" { $summary.Healthy++ }
                "serviceDegradation" { $summary.Advisory++ }
                "serviceInterruption" { $summary.Incident++ }
                default { $summary.Advisory++ }
            }
        }
        
        return [PSCustomObject]$summary
    }
    catch {
        Write-M365Log -Message "Failed to get health: $_" -Level Warning -Component "TenantMgmt"
        return $null
    }
}

function Get-M365TenantDomains {
    <#
    .SYNOPSIS
        Get all domains for a tenant
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId
    )
    
    if ($TenantId) {
        Switch-M365Tenant -TenantId $TenantId
    }
    
    $domains = Invoke-M365GraphRequest -Method GET -Uri "/domains"
    
    return $domains.value | Select-Object id, isDefault, isVerified, authenticationType, availabilityStatus
}

function Get-M365TenantSubscriptions {
    <#
    .SYNOPSIS
        Get subscription information for a tenant
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId
    )
    
    if ($TenantId) {
        Switch-M365Tenant -TenantId $TenantId
    }
    
    $skus = Invoke-M365GraphRequest -Method GET -Uri "/subscribedSkus"
    
    return $skus.value | Select-Object skuPartNumber, 
        @{N='Total';E={$_.prepaidUnits.enabled}},
        @{N='Assigned';E={$_.consumedUnits}},
        @{N='Available';E={$_.prepaidUnits.enabled - $_.consumedUnits}},
        appliesTo
}

Export-ModuleMember -Function @(
    'Get-M365TenantInfo',
    'Get-M365TenantHealth',
    'Get-M365TenantDomains',
    'Get-M365TenantSubscriptions'
)
