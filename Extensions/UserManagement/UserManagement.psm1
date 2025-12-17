<#
.SYNOPSIS
    User Management Extension
.DESCRIPTION
    Functions for managing users and groups across tenants.
#>

function Get-M365Users {
    <#
    .SYNOPSIS
        Get users from current tenant
    .PARAMETER Filter
        OData filter expression
    .PARAMETER All
        Get all users (may be slow for large tenants)
    .PARAMETER GuestsOnly
        Only return guest users
    .PARAMETER DisabledOnly
        Only return disabled users
    #>
    [CmdletBinding()]
    param(
        [string]$Filter,
        [switch]$All,
        [switch]$GuestsOnly,
        [switch]$DisabledOnly
    )
    
    $select = "id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedLicenses,signInActivity,createdDateTime,department,jobTitle"
    $uri = "/users?`$select=$select&`$top=100"
    
    # Build filter
    $filters = @()
    if ($GuestsOnly) {
        $filters += "userType eq 'Guest'"
    }
    if ($DisabledOnly) {
        $filters += "accountEnabled eq false"
    }
    if ($Filter) {
        $filters += $Filter
    }
    
    if ($filters.Count -gt 0) {
        $uri += "&`$filter=" + ($filters -join " and ")
    }
    
    if ($All) {
        return Get-M365GraphAllPages -Uri $uri
    }
    else {
        $response = Invoke-M365GraphRequest -Method GET -Uri $uri
        return $response.value
    }
}

function Get-M365User {
    <#
    .SYNOPSIS
        Get a specific user by ID or UPN
    .PARAMETER UserId
        User ID (GUID) or UPN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    
    $select = "id,displayName,userPrincipalName,mail,accountEnabled,userType,assignedLicenses,signInActivity,createdDateTime,department,jobTitle,manager,onPremisesSyncEnabled"
    
    return Invoke-M365GraphRequest -Method GET -Uri "/users/$UserId`?`$select=$select"
}

function Get-M365UserLicenses {
    <#
    .SYNOPSIS
        Get licenses assigned to a user
    .PARAMETER UserId
        User ID or UPN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    
    return Invoke-M365GraphRequest -Method GET -Uri "/users/$UserId/licenseDetails"
}

function Get-M365UserGroups {
    <#
    .SYNOPSIS
        Get groups a user is a member of
    .PARAMETER UserId
        User ID or UPN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    
    $groups = Invoke-M365GraphRequest -Method GET -Uri "/users/$UserId/memberOf"
    return $groups.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
}

function Get-M365UserSignInActivity {
    <#
    .SYNOPSIS
        Get sign-in activity for a user
    .PARAMETER UserId
        User ID or UPN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    
    $user = Invoke-M365GraphRequest -Method GET -Uri "/users/$UserId`?`$select=signInActivity"
    return $user.signInActivity
}

function Get-M365Groups {
    <#
    .SYNOPSIS
        Get groups from current tenant
    .PARAMETER Type
        Filter by group type (Security, M365, Distribution)
    .PARAMETER All
        Get all groups
    #>
    [CmdletBinding()]
    param(
        [ValidateSet("Security", "M365", "Distribution", "All")]
        [string]$Type = "All",
        
        [switch]$All
    )
    
    $select = "id,displayName,description,groupTypes,membershipRule,membershipRuleProcessingState,mail,mailEnabled,securityEnabled"
    $uri = "/groups?`$select=$select&`$top=100"
    
    # Filter by type
    switch ($Type) {
        "Security" {
            $uri += "&`$filter=securityEnabled eq true and mailEnabled eq false"
        }
        "M365" {
            $uri += "&`$filter=groupTypes/any(c:c eq 'Unified')"
        }
        "Distribution" {
            $uri += "&`$filter=mailEnabled eq true and securityEnabled eq false"
        }
    }
    
    if ($All) {
        return Get-M365GraphAllPages -Uri $uri
    }
    else {
        $response = Invoke-M365GraphRequest -Method GET -Uri $uri
        return $response.value
    }
}

function Get-M365GroupMembers {
    <#
    .SYNOPSIS
        Get members of a group
    .PARAMETER GroupId
        Group ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )
    
    return Get-M365GraphAllPages -Uri "/groups/$GroupId/members"
}

function Get-M365GroupOwners {
    <#
    .SYNOPSIS
        Get owners of a group
    .PARAMETER GroupId
        Group ID
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupId
    )
    
    $owners = Invoke-M365GraphRequest -Method GET -Uri "/groups/$GroupId/owners"
    return $owners.value
}

function Export-M365Users {
    <#
    .SYNOPSIS
        Export users to CSV or JSON
    .PARAMETER Path
        Output file path
    .PARAMETER Format
        Output format (CSV or JSON)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        
        [ValidateSet("CSV", "JSON")]
        [string]$Format = "CSV"
    )
    
    Write-M365Log -Message "Exporting users to: $Path" -Level Information -Component "UserMgmt"
    
    $users = Get-M365Users -All
    
    if ($Format -eq "CSV") {
        $users | Select-Object id, displayName, userPrincipalName, mail, accountEnabled, userType, department, jobTitle | 
            Export-Csv -Path $Path -NoTypeInformation
    }
    else {
        $users | ConvertTo-Json -Depth 10 | Set-Content $Path -Encoding UTF8
    }
    
    Write-M365Log -Message "Exported $($users.Count) users" -Level Information -Component "UserMgmt"
    
    return $Path
}

function Export-M365Groups {
    <#
    .SYNOPSIS
        Export groups to CSV or JSON
    .PARAMETER Path
        Output file path
    .PARAMETER Format
        Output format (CSV or JSON)
    .PARAMETER IncludeMembers
        Include member lists (JSON only)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        
        [ValidateSet("CSV", "JSON")]
        [string]$Format = "CSV",
        
        [switch]$IncludeMembers
    )
    
    Write-M365Log -Message "Exporting groups to: $Path" -Level Information -Component "UserMgmt"
    
    $groups = Get-M365Groups -All
    
    if ($IncludeMembers -and $Format -eq "JSON") {
        foreach ($group in $groups) {
            $members = Get-M365GroupMembers -GroupId $group.id
            Add-Member -InputObject $group -MemberType NoteProperty -Name "members" -Value $members -Force
        }
    }
    
    if ($Format -eq "CSV") {
        $groups | Select-Object id, displayName, description, mail, mailEnabled, securityEnabled | 
            Export-Csv -Path $Path -NoTypeInformation
    }
    else {
        $groups | ConvertTo-Json -Depth 10 | Set-Content $Path -Encoding UTF8
    }
    
    Write-M365Log -Message "Exported $($groups.Count) groups" -Level Information -Component "UserMgmt"
    
    return $Path
}

Export-ModuleMember -Function @(
    'Get-M365Users',
    'Get-M365User',
    'Get-M365UserLicenses',
    'Get-M365UserGroups',
    'Get-M365UserSignInActivity',
    'Get-M365Groups',
    'Get-M365GroupMembers',
    'Get-M365GroupOwners',
    'Export-M365Users',
    'Export-M365Groups'
)
