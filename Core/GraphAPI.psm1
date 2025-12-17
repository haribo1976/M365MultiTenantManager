<#
.SYNOPSIS
    Microsoft Graph API Module
.DESCRIPTION
    Provides wrapper functions for Microsoft Graph API calls
    with built-in retry logic, throttling handling, and pagination.
#>

#region Module Variables
$script:GraphBaseUrl = "https://graph.microsoft.com"
$script:GraphVersion = "v1.0"
$script:BetaEndpoints = @("secureScores", "conditionalAccess", "security")
#endregion

#region Private Functions
function Get-GraphHeaders {
    <#
    .SYNOPSIS
        Build headers for Graph API request
    #>
    $token = Get-M365AccessToken
    
    return @{
        "Authorization" = "Bearer $token"
        "Content-Type"  = "application/json"
        "ConsistencyLevel" = "eventual"
    }
}

function Should-UseBeta {
    <#
    .SYNOPSIS
        Determine if endpoint should use beta API
    #>
    param([string]$Uri)
    
    foreach ($endpoint in $script:BetaEndpoints) {
        if ($Uri -like "*$endpoint*") {
            return $true
        }
    }
    return $false
}

function Wait-WithBackoff {
    <#
    .SYNOPSIS
        Wait with exponential backoff
    #>
    param(
        [int]$Attempt,
        [int]$RetryAfter = 0
    )
    
    if ($RetryAfter -gt 0) {
        $waitTime = $RetryAfter
    }
    else {
        # Exponential backoff with jitter
        $baseWait = [Math]::Pow(2, $Attempt)
        $jitter = Get-Random -Minimum 0 -Maximum 1000
        $waitTime = $baseWait + ($jitter / 1000)
    }
    
    Write-M365Log -Message "Waiting $waitTime seconds before retry..." -Level Debug -Component "GraphAPI"
    Start-Sleep -Seconds $waitTime
}
#endregion

#region Public Functions
function Invoke-M365GraphRequest {
    <#
    .SYNOPSIS
        Make a Microsoft Graph API request with retry logic
    .PARAMETER Method
        HTTP method (GET, POST, PATCH, DELETE)
    .PARAMETER Uri
        API endpoint (relative to graph.microsoft.com)
    .PARAMETER Body
        Request body for POST/PATCH
    .PARAMETER Version
        API version (v1.0 or beta)
    .PARAMETER MaxRetries
        Maximum retry attempts for throttling
    .EXAMPLE
        Invoke-M365GraphRequest -Method GET -Uri "/users"
    .EXAMPLE
        Invoke-M365GraphRequest -Method POST -Uri "/users" -Body $userObject
    #>
    [CmdletBinding()]
    param(
        [ValidateSet("GET", "POST", "PATCH", "PUT", "DELETE")]
        [string]$Method = "GET",
        
        [Parameter(Mandatory)]
        [string]$Uri,
        
        [object]$Body,
        
        [ValidateSet("v1.0", "beta")]
        [string]$Version,
        
        [int]$MaxRetries = 3
    )
    
    # Determine version
    if ([string]::IsNullOrEmpty($Version)) {
        if (Should-UseBeta -Uri $Uri) {
            $Version = "beta"
        }
        else {
            $Version = Get-M365Setting -Name "GraphVersion" -Default "v1.0"
        }
    }
    
    # Build full URL
    if ($Uri.StartsWith("http")) {
        $fullUrl = $Uri
    }
    else {
        $Uri = $Uri.TrimStart("/")
        $fullUrl = "$($script:GraphBaseUrl)/$Version/$Uri"
    }
    
    $headers = Get-GraphHeaders
    
    for ($attempt = 0; $attempt -lt $MaxRetries; $attempt++) {
        try {
            Write-M365Log -Message "$Method $fullUrl" -Level Debug -Component "GraphAPI"
            
            $params = @{
                Uri         = $fullUrl
                Method      = $Method
                Headers     = $headers
                ErrorAction = "Stop"
            }
            
            if ($Body) {
                if ($Body -is [string]) {
                    $params.Body = $Body
                }
                else {
                    $params.Body = $Body | ConvertTo-Json -Depth 20
                }
            }
            
            $response = Invoke-RestMethod @params
            return $response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            if ($statusCode -eq 429) {
                # Throttled - get Retry-After header
                $retryAfter = 60
                try {
                    $retryAfter = [int]$_.Exception.Response.Headers["Retry-After"]
                }
                catch { }
                
                Write-M365Log -Message "Request throttled (429). Retry after: $retryAfter seconds" -Level Warning -Component "GraphAPI"
                
                if ($attempt -lt ($MaxRetries - 1)) {
                    Wait-WithBackoff -Attempt $attempt -RetryAfter $retryAfter
                    continue
                }
            }
            elseif ($statusCode -in @(500, 502, 503, 504)) {
                # Server error - retry with backoff
                Write-M365Log -Message "Server error ($statusCode). Retrying..." -Level Warning -Component "GraphAPI"
                
                if ($attempt -lt ($MaxRetries - 1)) {
                    Wait-WithBackoff -Attempt $attempt
                    continue
                }
            }
            
            # Log and throw
            Write-M365Log -Message "Graph API error: $($_.Exception.Message)" -Level Error -Component "GraphAPI"
            throw
        }
    }
    
    throw "Max retries exceeded for: $fullUrl"
}

function Get-M365GraphAllPages {
    <#
    .SYNOPSIS
        Get all pages of a Graph API response
    .PARAMETER Uri
        API endpoint
    .PARAMETER MaxPages
        Maximum pages to retrieve (0 = unlimited)
    .EXAMPLE
        Get-M365GraphAllPages -Uri "/users?`$top=100"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,
        
        [int]$MaxPages = 0
    )
    
    $allResults = @()
    $currentUri = $Uri
    $pageCount = 0
    
    do {
        $response = Invoke-M365GraphRequest -Method GET -Uri $currentUri
        
        if ($response.value) {
            $allResults += $response.value
        }
        else {
            $allResults += $response
        }
        
        $currentUri = $response.'@odata.nextLink'
        $pageCount++
        
        if ($MaxPages -gt 0 -and $pageCount -ge $MaxPages) {
            Write-M365Log -Message "Reached max pages limit: $MaxPages" -Level Debug -Component "GraphAPI"
            break
        }
        
    } while ($currentUri)
    
    Write-M365Log -Message "Retrieved $($allResults.Count) items in $pageCount pages" -Level Debug -Component "GraphAPI"
    
    return $allResults
}

function Invoke-M365GraphBatch {
    <#
    .SYNOPSIS
        Execute batch Graph API requests (max 20 per batch)
    .PARAMETER Requests
        Array of request objects with id, method, and url
    .EXAMPLE
        Invoke-M365GraphBatch -Requests @(
            @{ id = "1"; method = "GET"; url = "/users?`$top=5" }
            @{ id = "2"; method = "GET"; url = "/groups?`$top=5" }
        )
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Requests
    )
    
    if ($Requests.Count -gt 20) {
        Write-M365Log -Message "Batch limited to 20 requests. Splitting..." -Level Warning -Component "GraphAPI"
        
        $results = @()
        for ($i = 0; $i -lt $Requests.Count; $i += 20) {
            $batch = $Requests[$i..([Math]::Min($i + 19, $Requests.Count - 1))]
            $results += Invoke-M365GraphBatch -Requests $batch
        }
        return $results
    }
    
    $batchBody = @{
        requests = $Requests
    }
    
    $response = Invoke-M365GraphRequest -Method POST -Uri '/$batch' -Body $batchBody
    
    return $response.responses
}

function Test-M365GraphConnection {
    <#
    .SYNOPSIS
        Test Graph API connectivity
    #>
    try {
        $org = Invoke-M365GraphRequest -Method GET -Uri "/organization"
        return $true
    }
    catch {
        return $false
    }
}

function Get-M365GraphRateLimitInfo {
    <#
    .SYNOPSIS
        Get rate limit info from last request (if available)
    #>
    # This would require storing headers from last response
    # Placeholder for future implementation
    return [PSCustomObject]@{
        RemainingCalls = "Unknown"
        ResetTime      = "Unknown"
    }
}
#endregion

Export-ModuleMember -Function @(
    'Invoke-M365GraphRequest',
    'Get-M365GraphAllPages',
    'Invoke-M365GraphBatch',
    'Test-M365GraphConnection',
    'Get-M365GraphRateLimitInfo'
)
