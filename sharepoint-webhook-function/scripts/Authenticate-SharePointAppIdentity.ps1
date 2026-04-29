#Requires -Version 5.1
<#
.SYNOPSIS
    Authenticates an Azure AD App Identity to SharePoint Online.

.DESCRIPTION
    Supports three authentication methods:
      - ServicePrincipal  : Client Credentials (certificate or client secret)
      - ManagedIdentity   : Azure Managed Identity (System or User-assigned)
      - Interactive        : Device code / browser-based interactive login

    Returns a Bearer access token that can be used with the SharePoint REST API
    or Microsoft Graph.

.PARAMETER AuthMethod
    Authentication method to use. Valid values: ServicePrincipal, ManagedIdentity, Interactive.

.PARAMETER TenantId
    Azure AD tenant ID (GUID or domain). Required for ServicePrincipal and Interactive methods.

.PARAMETER ClientId
    Azure AD application (client) ID. Required for ServicePrincipal; optional for ManagedIdentity
    (used to specify a user-assigned identity).

.PARAMETER ClientSecret
    Client secret string. Used when CertificatePath is not provided and AuthMethod is ServicePrincipal.

.PARAMETER CertificatePath
    Path to a PFX certificate file. Used instead of ClientSecret when AuthMethod is ServicePrincipal.

.PARAMETER CertificatePassword
    Password for the PFX certificate. Required when CertificatePath is provided.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate already installed in the local certificate store.
    Alternative to CertificatePath.

.PARAMETER SharePointSiteUrl
    Full URL of the target SharePoint site collection (e.g. https://contoso.sharepoint.com/sites/mysite).

.PARAMETER OutputTokenToFile
    Optional path to write the access token to a file (useful for piping to other scripts).

.EXAMPLE
    # Service Principal with client secret
    .\Authenticate-SharePointAppIdentity.ps1 `
        -AuthMethod ServicePrincipal `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -ClientSecret "your-client-secret" `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite"

.EXAMPLE
    # Service Principal with certificate file
    .\Authenticate-SharePointAppIdentity.ps1 `
        -AuthMethod ServicePrincipal `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -CertificatePath "C:\certs\myapp.pfx" `
        -CertificatePassword (Read-Host -AsSecureString "Cert password") `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite"

.EXAMPLE
    # System-assigned Managed Identity (running inside Azure)
    .\Authenticate-SharePointAppIdentity.ps1 `
        -AuthMethod ManagedIdentity `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite"

.EXAMPLE
    # Interactive / device code (useful for admin tasks)
    .\Authenticate-SharePointAppIdentity.ps1 `
        -AuthMethod Interactive `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("ServicePrincipal", "ManagedIdentity", "Interactive")]
    [string]$AuthMethod,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,

    [Parameter(Mandatory = $false)]
    [securestring]$CertificatePassword,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [string]$SharePointSiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputTokenToFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Helpers ──────────────────────────────────────────────────────────────

function Write-Banner {
    param([string]$Title)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Write-Step {
    param([string]$Message)
    Write-Host "`n[*] $Message" -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Message)
    Write-Host "[-] $Message" -ForegroundColor Red
}

# Extract the SharePoint tenant root from a site URL
function Get-TenantRootUrl {
    param([string]$SiteUrl)
    $uri = [System.Uri]$SiteUrl
    return "$($uri.Scheme)://$($uri.Host)"
}

# Derive the SharePoint resource audience (tenant root)
function Get-SharePointResource {
    param([string]$SiteUrl)
    $tenantRoot = Get-TenantRootUrl -SiteUrl $SiteUrl
    return "$tenantRoot/"
}

#endregion

#region ── Authentication Methods ───────────────────────────────────────────────

function Get-TokenWithClientSecret {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Resource
    )

    Write-Step "Acquiring token via Client Credentials (client secret)..."

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = $Resource
    }

    try {
        $response = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        $statusCode = $_.Exception.Response?.StatusCode.value__
        $errorBody  = $_.ErrorDetails?.Message
        Write-Fail "Token request failed (HTTP $statusCode): $errorBody"
        throw
    }
}

function Get-TokenWithCertificate {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Resource
    )

    Write-Step "Acquiring token via Client Credentials (certificate)..."

    # Build JWT header
    $now       = [DateTimeOffset]::UtcNow
    $nbf       = $now.ToUnixTimeSeconds()
    $exp       = $now.AddMinutes(10).ToUnixTimeSeconds()
    $jti       = [System.Guid]::NewGuid().ToString()
    $audience  = "https://login.microsoftonline.com/$TenantId/oauth2/token"
    $thumbBytes = [System.Convert]::FromHexString($Certificate.Thumbprint)
    $x5t       = [System.Convert]::ToBase64String($thumbBytes)

    $header  = [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{ alg = "RS256"; typ = "JWT"; x5t = $x5t } -Compress)
        )
    ).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $payload = [System.Convert]::ToBase64String(
        [System.Text.Encoding]::UTF8.GetBytes(
            (ConvertTo-Json @{
                aud = $audience
                iss = $ClientId
                sub = $ClientId
                jti = $jti
                nbf = $nbf
                exp = $exp
            } -Compress)
        )
    ).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $signingInput = "$header.$payload"
    $rsa           = $Certificate.GetRSAPrivateKey()
    $sigBytes      = $rsa.SignData(
        [System.Text.Encoding]::UTF8.GetBytes($signingInput),
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $signature = [System.Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $clientAssertion = "$signingInput.$signature"

    $body = @{
        grant_type             = "client_credentials"
        client_id              = $ClientId
        client_assertion_type  = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion       = $clientAssertion
        resource               = $Resource
    }

    try {
        $response = Invoke-RestMethod -Uri $audience -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        return $response.access_token
    }
    catch {
        $statusCode = $_.Exception.Response?.StatusCode.value__
        $errorBody  = $_.ErrorDetails?.Message
        Write-Fail "Token request failed (HTTP $statusCode): $errorBody"
        throw
    }
}

function Get-TokenWithManagedIdentity {
    param(
        [string]$ClientId,   # Optional: user-assigned identity client ID
        [string]$Resource
    )

    Write-Step "Acquiring token via Managed Identity..."

    # IMDS endpoint available on all Azure hosts
    $imdsUrl = "http://169.254.169.254/metadata/identity/oauth2/token"
    $params  = "api-version=2019-08-01&resource=$([System.Uri]::EscapeDataString($Resource))"

    if ($ClientId) {
        $params += "&client_id=$([System.Uri]::EscapeDataString($ClientId))"
        Write-Host "  Using user-assigned Managed Identity: $ClientId" -ForegroundColor DarkGray
    }
    else {
        Write-Host "  Using system-assigned Managed Identity" -ForegroundColor DarkGray
    }

    $headers = @{ Metadata = "true" }

    try {
        $response = Invoke-RestMethod -Uri "$imdsUrl`?$params" -Method Get -Headers $headers -TimeoutSec 10
        return $response.access_token
    }
    catch {
        $statusCode = $_.Exception.Response?.StatusCode.value__
        $errorBody  = $_.ErrorDetails?.Message
        Write-Fail "Managed Identity token request failed (HTTP $statusCode): $errorBody"
        Write-Host "  Make sure this script is running inside an Azure resource with Managed Identity enabled." -ForegroundColor DarkYellow
        throw
    }
}

function Get-TokenInteractive {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Resource
    )

    Write-Step "Acquiring token via Device Code (interactive)..."

    # Device code flow
    $deviceCodeEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/devicecode"
    $tokenEndpoint      = "https://login.microsoftonline.com/$TenantId/oauth2/token"

    $dcResponse = Invoke-RestMethod -Uri $deviceCodeEndpoint -Method Post -Body @{
        client_id = $ClientId
        resource  = $Resource
    } -ContentType "application/x-www-form-urlencoded"

    Write-Host ""
    Write-Host $dcResponse.message -ForegroundColor Cyan
    Write-Host ""

    # Poll for completion
    $interval    = [int]$dcResponse.interval
    $expiresIn   = [int]$dcResponse.expires_in
    $deviceCode  = $dcResponse.device_code
    $startTime   = Get-Date

    while ($true) {
        Start-Sleep -Seconds $interval

        if ((Get-Date) -gt $startTime.AddSeconds($expiresIn)) {
            throw "Device code expired. Please re-run the script and authenticate within the time limit."
        }

        try {
            $tokenResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body @{
                grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                client_id   = $ClientId
                device_code = $deviceCode
            } -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

            return $tokenResponse.access_token
        }
        catch {
            $errorBody = $_.ErrorDetails?.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($errorBody?.error -eq "authorization_pending") {
                Write-Host "  Waiting for user authentication..." -ForegroundColor DarkGray
            }
            elseif ($errorBody?.error -eq "slow_down") {
                $interval += 5
            }
            else {
                throw
            }
        }
    }
}

#endregion

#region ── Main ─────────────────────────────────────────────────────────────────

Write-Banner "SharePoint App Identity Authentication"
Write-Host "  Auth Method : $AuthMethod"
Write-Host "  Site URL    : $SharePointSiteUrl"

$resource = Get-SharePointResource -SiteUrl $SharePointSiteUrl
Write-Host "  Resource    : $resource"

$accessToken = $null

switch ($AuthMethod) {

    "ServicePrincipal" {
        if (-not $TenantId) { throw "TenantId is required for ServicePrincipal authentication." }
        if (-not $ClientId)  { throw "ClientId is required for ServicePrincipal authentication." }

        # Certificate takes priority over client secret
        if ($CertificatePath -or $CertificateThumbprint) {

            $cert = $null
            if ($CertificatePath) {
                Write-Step "Loading certificate from file: $CertificatePath"
                if (-not (Test-Path $CertificatePath)) {
                    throw "Certificate file not found: $CertificatePath"
                }
                if ($CertificatePassword) {
                    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                        $CertificatePath,
                        $CertificatePassword,
                        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet -bor
                        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
                    )
                }
                else {
                    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath)
                }
            }
            else {
                Write-Step "Loading certificate from store (thumbprint: $CertificateThumbprint)..."
                $cert = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if (-not $cert) {
                    $cert = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                }
                if (-not $cert) {
                    throw "Certificate with thumbprint '$CertificateThumbprint' not found in CurrentUser or LocalMachine store."
                }
            }

            Write-Success "Certificate loaded: Subject=$($cert.Subject), Thumbprint=$($cert.Thumbprint)"
            $accessToken = Get-TokenWithCertificate -TenantId $TenantId -ClientId $ClientId -Certificate $cert -Resource $resource
        }
        elseif ($ClientSecret) {
            $accessToken = Get-TokenWithClientSecret -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -Resource $resource
        }
        else {
            throw "For ServicePrincipal authentication, provide either -ClientSecret or -CertificatePath / -CertificateThumbprint."
        }
    }

    "ManagedIdentity" {
        $accessToken = Get-TokenWithManagedIdentity -ClientId $ClientId -Resource $resource
    }

    "Interactive" {
        if (-not $TenantId) { throw "TenantId is required for Interactive authentication." }
        if (-not $ClientId)  { throw "ClientId is required for Interactive authentication." }
        $accessToken = Get-TokenInteractive -TenantId $TenantId -ClientId $ClientId -Resource $resource
    }
}

if (-not $accessToken) {
    throw "Failed to obtain access token."
}

#region ── Validate token by calling SharePoint ──────────────────────────────────
Write-Step "Validating token against SharePoint..."

$testUrl = "$($SharePointSiteUrl.TrimEnd('/'))/_api/web?`$select=Title,Url"
try {
    $spResponse = Invoke-RestMethod -Uri $testUrl -Method Get -Headers @{
        Authorization = "Bearer $accessToken"
        Accept        = "application/json;odata=nometadata"
    }
    Write-Success "Token validated! Connected to site: '$($spResponse.Title)' ($($spResponse.Url))"
}
catch {
    $statusCode = $_.Exception.Response?.StatusCode.value__
    Write-Fail "SharePoint validation call failed (HTTP $statusCode). The token may lack required permissions."
    Write-Host "  Check that the app has 'Sites.Selected' or 'Sites.Manage.All' SharePoint permission." -ForegroundColor DarkYellow
}
#endregion

#region ── Output ────────────────────────────────────────────────────────────────
Write-Banner "Authentication Result"

# Show truncated token (never log full tokens in production)
$preview = $accessToken.Substring(0, [Math]::Min(40, $accessToken.Length)) + "..."
Write-Host "  Token preview : $preview" -ForegroundColor DarkGray
Write-Host "  Token length  : $($accessToken.Length) characters" -ForegroundColor DarkGray

if ($OutputTokenToFile) {
    $accessToken | Out-File -FilePath $OutputTokenToFile -Encoding UTF8 -NoNewline
    Write-Success "Token written to: $OutputTokenToFile"
}

# Export to pipeline / calling script via global variable
$global:SharePointAccessToken = $accessToken
Write-Success "Access token stored in `$global:SharePointAccessToken"

Write-Host ""
Write-Host "Usage example:" -ForegroundColor Cyan
Write-Host '  $headers = @{ Authorization = "Bearer $global:SharePointAccessToken"; Accept = "application/json;odata=nometadata" }' -ForegroundColor DarkGray
Write-Host "  Invoke-RestMethod -Uri `"$SharePointSiteUrl/_api/web/lists`" -Headers `$headers" -ForegroundColor DarkGray
Write-Host ""

return $accessToken
#endregion
