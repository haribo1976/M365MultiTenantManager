<#
.SYNOPSIS
    Security Management Extension
.DESCRIPTION
    Functions for managing Conditional Access policies and Secure Score.
#>

#region Conditional Access

function Get-M365ConditionalAccessPolicies {
    <#
    .SYNOPSIS
        Get all Conditional Access policies
    #>
    [CmdletBinding()]
    param()
    
    $policies = Invoke-M365GraphRequest -Method GET -Uri "/identity/conditionalAccess/policies" -Version "beta"
    
    return $policies.value | ForEach-Object {
        [PSCustomObject]@{
            Id              = $_.id
            DisplayName     = $_.displayName
            State           = $_.state
            CreatedDateTime = $_.createdDateTime
            ModifiedDateTime = $_.modifiedDateTime
            Conditions      = $_.conditions
            GrantControls   = $_.grantControls
            SessionControls = $_.sessionControls
        }
    }
}

function Get-M365ConditionalAccessPolicy {
    <#
    .SYNOPSIS
        Get a specific Conditional Access policy
    .PARAMETER PolicyId
        Policy ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PolicyId
    )
    
    return Invoke-M365GraphRequest -Method GET -Uri "/identity/conditionalAccess/policies/$PolicyId" -Version "beta"
}

function Export-M365ConditionalAccessPolicy {
    <#
    .SYNOPSIS
        Export Conditional Access policy to JSON
    .PARAMETER PolicyId
        Policy ID (or 'All' for all policies)
    .PARAMETER Path
        Output path
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$PolicyId,
        
        [Parameter(Mandatory)]
        [string]$Path
    )
    
    Write-M365Log -Message "Exporting CA policy: $PolicyId" -Level Information -Component "SecurityMgmt"
    
    if ($PolicyId -eq "All") {
        $policies = Get-M365ConditionalAccessPolicies
        
        # Create directory for all policies
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
        
        foreach ($policy in $policies) {
            $fileName = "$($policy.DisplayName -replace '[^\w\s-]', '')_$($policy.Id).json"
            $filePath = Join-Path $Path $fileName
            
            # Get full policy details
            $fullPolicy = Get-M365ConditionalAccessPolicy -PolicyId $policy.Id
            $fullPolicy | ConvertTo-Json -Depth 20 | Set-Content $filePath -Encoding UTF8
        }
        
        Write-M365Log -Message "Exported $($policies.Count) policies to: $Path" -Level Information -Component "SecurityMgmt"
    }
    else {
        $policy = Get-M365ConditionalAccessPolicy -PolicyId $PolicyId
        $policy | ConvertTo-Json -Depth 20 | Set-Content $Path -Encoding UTF8
        Write-M365Log -Message "Exported policy to: $Path" -Level Information -Component "SecurityMgmt"
    }
    
    return $Path
}

function Import-M365ConditionalAccessPolicy {
    <#
    .SYNOPSIS
        Import Conditional Access policy from JSON
    .PARAMETER Path
        Path to JSON file
    .PARAMETER State
        Override policy state (enabled, disabled, enabledForReportingButNotEnforced)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        
        [ValidateSet("enabled", "disabled", "enabledForReportingButNotEnforced")]
        [string]$State = "disabled"
    )
    
    if (-not (Test-Path $Path)) {
        throw "File not found: $Path"
    }
    
    $policy = Get-Content $Path -Raw | ConvertFrom-Json
    
    # Remove read-only properties
    $importPolicy = @{
        displayName     = $policy.displayName
        state           = $State
        conditions      = $policy.conditions
        grantControls   = $policy.grantControls
    }
    
    if ($policy.sessionControls) {
        $importPolicy.sessionControls = $policy.sessionControls
    }
    
    Write-M365Log -Message "Importing CA policy: $($policy.displayName)" -Level Information -Component "SecurityMgmt"
    
    $result = Invoke-M365GraphRequest -Method POST -Uri "/identity/conditionalAccess/policies" -Body $importPolicy -Version "beta"
    
    Write-M365Log -Message "Policy created with ID: $($result.id)" -Level Information -Component "SecurityMgmt"
    
    return $result
}

function Compare-M365ConditionalAccessPolicies {
    <#
    .SYNOPSIS
        Compare CA policies between two tenants
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
    
    # Get source policies
    Switch-M365Tenant -TenantId $SourceTenantId
    $sourcePolicies = Get-M365ConditionalAccessPolicies
    
    # Get target policies
    Switch-M365Tenant -TenantId $TargetTenantId
    $targetPolicies = Get-M365ConditionalAccessPolicies
    
    $comparison = @()
    
    # Check source policies against target
    foreach ($source in $sourcePolicies) {
        $match = $targetPolicies | Where-Object { $_.DisplayName -eq $source.DisplayName }
        
        if ($match) {
            $status = "Match"
            if ($source.State -ne $match.State) {
                $status = "StateDifferent"
            }
        }
        else {
            $status = "SourceOnly"
            $match = $null
        }
        
        $comparison += [PSCustomObject]@{
            PolicyName    = $source.DisplayName
            Status        = $status
            SourceId      = $source.Id
            SourceState   = $source.State
            TargetId      = $match.Id
            TargetState   = $match.State
        }
    }
    
    # Find target-only policies
    foreach ($target in $targetPolicies) {
        $exists = $comparison | Where-Object { $_.PolicyName -eq $target.DisplayName }
        if (-not $exists) {
            $comparison += [PSCustomObject]@{
                PolicyName    = $target.DisplayName
                Status        = "TargetOnly"
                SourceId      = $null
                SourceState   = $null
                TargetId      = $target.Id
                TargetState   = $target.State
            }
        }
    }
    
    return $comparison
}

#endregion

#region Secure Score

function Get-M365SecureScore {
    <#
    .SYNOPSIS
        Get current Secure Score
    #>
    [CmdletBinding()]
    param()
    
    $scores = Invoke-M365GraphRequest -Method GET -Uri "/security/secureScores?`$top=1" -Version "beta"
    
    if ($scores.value.Count -eq 0) {
        return $null
    }
    
    $current = $scores.value[0]
    
    return [PSCustomObject]@{
        CurrentScore  = $current.currentScore
        MaxScore      = $current.maxScore
        Percentage    = [math]::Round(($current.currentScore / $current.maxScore) * 100, 1)
        CreatedDate   = $current.createdDateTime
        ControlScores = $current.controlScores
        AverageComparativeScores = $current.averageComparativeScores
    }
}

function Get-M365SecureScoreControlProfiles {
    <#
    .SYNOPSIS
        Get Secure Score control profiles (improvement actions)
    #>
    [CmdletBinding()]
    param()
    
    $controls = Get-M365GraphAllPages -Uri "/security/secureScoreControlProfiles" -MaxPages 5
    
    return $controls | ForEach-Object {
        [PSCustomObject]@{
            Id                = $_.id
            Title             = $_.title
            MaxScore          = $_.maxScore
            Tier              = $_.tier
            ImplementationCost = $_.implementationCost
            UserImpact        = $_.userImpact
            Threats           = $_.threats
            ActionType        = $_.actionType
            Deprecated        = $_.deprecated
            Remediation       = $_.remediation
            RemediationImpact = $_.remediationImpact
        }
    }
}

function Compare-M365SecureScores {
    <#
    .SYNOPSIS
        Compare Secure Scores across all registered tenants
    #>
    [CmdletBinding()]
    param()
    
    $tenants = Get-M365RegisteredTenants
    $results = @()
    
    foreach ($tenant in $tenants) {
        Write-M365Log -Message "Getting Secure Score for: $($tenant.friendlyName)" -Level Debug -Component "SecurityMgmt"
        
        try {
            Switch-M365Tenant -TenantId $tenant.id
            $score = Get-M365SecureScore
            
            $results += [PSCustomObject]@{
                TenantId     = $tenant.id
                TenantName   = $tenant.friendlyName
                CurrentScore = $score.CurrentScore
                MaxScore     = $score.MaxScore
                Percentage   = $score.Percentage
                Status       = if ($score.Percentage -ge 80) { "Good" } 
                              elseif ($score.Percentage -ge 60) { "Fair" } 
                              else { "NeedsImprovement" }
            }
        }
        catch {
            Write-M365Log -Message "Failed to get Secure Score for $($tenant.friendlyName): $_" -Level Warning -Component "SecurityMgmt"
            
            $results += [PSCustomObject]@{
                TenantId     = $tenant.id
                TenantName   = $tenant.friendlyName
                CurrentScore = $null
                MaxScore     = $null
                Percentage   = $null
                Status       = "Error"
            }
        }
    }
    
    return $results | Sort-Object -Property Percentage -Descending
}

#endregion

Export-ModuleMember -Function @(
    'Get-M365ConditionalAccessPolicies',
    'Get-M365ConditionalAccessPolicy',
    'Export-M365ConditionalAccessPolicy',
    'Import-M365ConditionalAccessPolicy',
    'Compare-M365ConditionalAccessPolicies',
    'Get-M365SecureScore',
    'Get-M365SecureScoreControlProfiles',
    'Compare-M365SecureScores'
)
