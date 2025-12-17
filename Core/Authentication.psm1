<#
.SYNOPSIS
    Authentication Module
.DESCRIPTION
    Handles Microsoft Graph authentication using MSAL.
    Supports interactive, device code, and client credential flows.
#>

#region Module Variables
$script:CurrentToken = $null
$script:CurrentTenantId = $null
$script:CurrentAccount = $null
$script:TokenCache = @{}
$script:MsalAvailable = $false
#endregion

#region Private Functions
function Initialize-MSAL {
    <#
    .SYNOPSIS
        Initialize MSAL library
    #>
    
    # Try MSAL.PS module first
    if (Get-Module -ListAvailable -Name MSAL.PS) {
        Import-Module MSAL.PS -ErrorAction SilentlyContinue
        if (Get-Command Get-MsalToken -ErrorAction SilentlyContinue) {
            $script:MsalAvailable = $true
            Write-M365Log -Message "MSAL.PS module loaded" -Level Debug
            return
        }
    }
    
    # Try Microsoft.Identity.Client from Az module
    $azModule = Get-Module -ListAvailable -Name Az.Accounts | Select-Object -First 1
    if ($azModule) {
        $msalPath = Join-Path (Split-Path $azModule.Path -Parent) "Microsoft.Identity.Client.dll"
        if (Test-Path $msalPath) {
            Add-Type -Path $msalPath
            $script:MsalAvailable = $true
            Write-M365Log -Message "MSAL loaded from Az module" -Level Debug
            return
        }
    }
    
    # Try local Lib folder
    $localMsal = Join-Path $PSScriptRoot "..\Lib\MSAL\Microsoft.Identity.Client.dll"
    if (Test-Path $localMsal) {
        Add-Type -Path $localMsal
        $script:MsalAvailable = $true
        Write-M365Log -Message "MSAL loaded from local Lib folder" -Level Debug
        return
    }
    
    Write-M365Log -Message "MSAL not available. Install MSAL.PS module: Install-Module MSAL.PS" -Level Warning
}

function Get-GraphScopes {
    <#
    .SYNOPSIS
        Get required Graph API scopes
    #>
    param(
        [switch]$ReadOnly
    )
    
    $scopes = @(
        "https://graph.microsoft.com/User.Read.All",
        "https://graph.microsoft.com/Group.Read.All",
        "https://graph.microsoft.com/Directory.Read.All",
        "https://graph.microsoft.com/Organization.Read.All",
        "https://graph.microsoft.com/Policy.Read.All",
        "https://graph.microsoft.com/SecurityEvents.Read.All",
        "https://graph.microsoft.com/AuditLog.Read.All"
    )
    
    if (-not $ReadOnly) {
        $scopes += @(
            "https://graph.microsoft.com/Policy.ReadWrite.ConditionalAccess",
            "https://graph.microsoft.com/User.ReadWrite.All",
            "https://graph.microsoft.com/Group.ReadWrite.All"
        )
    }
    
    return $scopes
}
#endregion

#region Public Functions
function Connect-M365Tenant {
    <#
    .SYNOPSIS
        Connect to a Microsoft 365 tenant
    .PARAMETER TenantId
        Tenant ID (GUID) or domain name
    .PARAMETER AppId
        Application ID for authentication
    .PARAMETER Secret
        Client secret for app authentication
    .PARAMETER CertThumbprint
        Certificate thumbprint for app authentication
    .PARAMETER UseDeviceCode
        Use device code flow (for headless scenarios)
    .PARAMETER ReadOnly
        Request only read permissions
    .EXAMPLE
        Connect-M365Tenant -TenantId "contoso.onmicrosoft.com"
    .EXAMPLE
        Connect-M365Tenant -TenantId "xxx-xxx" -AppId "yyy" -Secret "zzz"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,
        
        [string]$AppId,
        
        [string]$Secret,
        
        [string]$CertThumbprint,
        
        [switch]$UseDeviceCode,
        
        [switch]$ReadOnly
    )
    
    # Initialize MSAL if needed
    if (-not $script:MsalAvailable) {
        Initialize-MSAL
    }
    
    # Use default app if not specified
    if ([string]::IsNullOrEmpty($AppId)) {
        $AppId = Get-M365Setting -Name "DefaultAppId"
    }
    
    $scopes = Get-GraphScopes -ReadOnly:$ReadOnly
    
    Write-M365Log -Message "Connecting to tenant: $TenantId" -Level Information -Component "Auth"
    
    try {
        if ($Secret) {
            # Client credentials flow
            Write-M365Log -Message "Using client credentials flow" -Level Debug -Component "Auth"
            
            if (Get-Command Get-MsalToken -ErrorAction SilentlyContinue) {
                $secureSecret = ConvertTo-SecureString $Secret -AsPlainText -Force
                $token = Get-MsalToken -ClientId $AppId -TenantId $TenantId -ClientSecret $secureSecret
            }
            else {
                # Direct REST call for client credentials
                $body = @{
                    client_id     = $AppId
                    client_secret = $Secret
                    scope         = "https://graph.microsoft.com/.default"
                    grant_type    = "client_credentials"
                }
                
                $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                $response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
                
                $token = [PSCustomObject]@{
                    AccessToken = $response.access_token
                    ExpiresOn   = (Get-Date).AddSeconds($response.expires_in)
                }
            }
        }
        elseif ($CertThumbprint) {
            # Certificate-based auth
            Write-M365Log -Message "Using certificate authentication" -Level Debug -Component "Auth"
            
            $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction SilentlyContinue
            if (-not $cert) {
                $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertThumbprint" -ErrorAction SilentlyContinue
            }
            
            if (-not $cert) {
                throw "Certificate not found: $CertThumbprint"
            }
            
            if (Get-Command Get-MsalToken -ErrorAction SilentlyContinue) {
                $token = Get-MsalToken -ClientId $AppId -TenantId $TenantId -ClientCertificate $cert
            }
            else {
                throw "Certificate auth requires MSAL.PS module"
            }
        }
        elseif ($UseDeviceCode) {
            # Device code flow
            Write-M365Log -Message "Using device code flow" -Level Debug -Component "Auth"
            
            if (Get-Command Get-MsalToken -ErrorAction SilentlyContinue) {
                $token = Get-MsalToken -ClientId $AppId -TenantId $TenantId -Scopes $scopes -DeviceCode
            }
            else {
                throw "Device code flow requires MSAL.PS module"
            }
        }
        else {
            # Interactive flow
            Write-M365Log -Message "Using interactive authentication" -Level Debug -Component "Auth"
            
            # Force import MSAL.PS
            Import-Module MSAL.PS -Force -ErrorAction SilentlyContinue
            if (Get-Command Get-MsalToken -ErrorAction SilentlyContinue) {
                $token = Get-MsalToken -ClientId $AppId -TenantId $TenantId -Scopes $scopes -Interactive
            }
            else {
                # Fallback to simple OAuth flow
                throw "Interactive auth requires MSAL.PS module. Install with: Install-Module MSAL.PS"
            }
        }
        
        # Store token
        $script:CurrentToken = $token
        $script:CurrentTenantId = $TenantId
        $script:TokenCache[$TenantId] = $token
        
        # Get account info if available
        if ($token.Account) {
            $script:CurrentAccount = $token.Account
        }
        
        Write-M365Log -Message "Successfully connected to tenant: $TenantId" -Level Information -Component "Auth"
        
        return $token
    }
    catch {
        Write-M365Log -Message "Authentication failed: $_" -Level Error -Component "Auth"
        throw
    }
}

function Disconnect-M365Tenant {
    <#
    .SYNOPSIS
        Disconnect from current tenant
    .PARAMETER All
        Clear all cached tokens
    #>
    [CmdletBinding()]
    param(
        [switch]$All
    )
    
    if ($All) {
        $script:TokenCache = @{}
        Write-M365Log -Message "Cleared all cached tokens" -Level Information -Component "Auth"
    }
    elseif ($script:CurrentTenantId) {
        $script:TokenCache.Remove($script:CurrentTenantId)
        Write-M365Log -Message "Disconnected from tenant: $($script:CurrentTenantId)" -Level Information -Component "Auth"
    }
    
    $script:CurrentToken = $null
    $script:CurrentTenantId = $null
    $script:CurrentAccount = $null
}

function Switch-M365Tenant {
    <#
    .SYNOPSIS
        Switch to a different tenant (uses cached token if available)
    .PARAMETER TenantId
        Tenant ID to switch to
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    
    # Check cache first
    if ($script:TokenCache.ContainsKey($TenantId)) {
        $cachedToken = $script:TokenCache[$TenantId]
        
        # Check if token is still valid
        if ($cachedToken.ExpiresOn -gt (Get-Date).AddMinutes(5)) {
            $script:CurrentToken = $cachedToken
            $script:CurrentTenantId = $TenantId
            Write-M365Log -Message "Switched to tenant (cached): $TenantId" -Level Information -Component "Auth"
            return $cachedToken
        }
    }
    
    # Need to re-authenticate
    Write-M365Log -Message "Token expired or not cached, re-authenticating" -Level Debug -Component "Auth"
    return Connect-M365Tenant -TenantId $TenantId
}

function Get-M365AuthContext {
    <#
    .SYNOPSIS
        Get current authentication context
    #>
    return [PSCustomObject]@{
        TenantId     = $script:CurrentTenantId
        Account      = $script:CurrentAccount
        IsConnected  = ($null -ne $script:CurrentToken)
        TokenExpiry  = $script:CurrentToken.ExpiresOn
        CachedTenants = $script:TokenCache.Keys
    }
}

function Get-M365AccessToken {
    <#
    .SYNOPSIS
        Get current access token for API calls
    #>
    if (-not $script:CurrentToken) {
        throw "Not connected. Use Connect-M365Tenant first."
    }
    
    # Check if token needs refresh
    if ($script:CurrentToken.ExpiresOn -lt (Get-Date).AddMinutes(5)) {
        Write-M365Log -Message "Token expiring soon, refreshing..." -Level Debug -Component "Auth"
        $script:CurrentToken = Connect-M365Tenant -TenantId $script:CurrentTenantId
    }
    
    return $script:CurrentToken.AccessToken
}

function Test-M365Connection {
    <#
    .SYNOPSIS
        Test if currently connected and token is valid
    #>
    if (-not $script:CurrentToken) {
        return $false
    }
    
    return $script:CurrentToken.ExpiresOn -gt (Get-Date)
}
#endregion

#region Module Initialization
Initialize-MSAL
#endregion

Export-ModuleMember -Function @(
    'Connect-M365Tenant',
    'Disconnect-M365Tenant',
    'Switch-M365Tenant',
    'Get-M365AuthContext',
    'Get-M365AccessToken',
    'Test-M365Connection'
)
