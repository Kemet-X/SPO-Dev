#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive testing and troubleshooting suite for SharePoint webhook
    authentication on Azure Functions.

.DESCRIPTION
    Runs a series of connectivity, authentication, and permission checks to help
    diagnose issues with a SharePoint webhook Azure Function setup.

    Test suites included:
      - Network / HTTPS connectivity
      - Azure AD token acquisition (service principal and / or managed identity)
      - SharePoint REST API access (site, list, subscriptions)
      - Azure Function endpoint validation (validation token handshake)
      - Webhook subscription status

.PARAMETER TenantId
    Azure AD tenant ID.

.PARAMETER ClientId
    Azure AD application (client) ID for Service Principal tests.

.PARAMETER ClientSecret
    Client secret for Service Principal tests (alternative to certificate).

.PARAMETER CertificatePath
    Path to a PFX certificate for Service Principal tests.

.PARAMETER CertificatePassword
    Password for the PFX certificate.

.PARAMETER SharePointSiteUrl
    Full URL of the SharePoint site collection.

.PARAMETER SharePointListName
    Display name of the SharePoint list to test webhook subscriptions on.

.PARAMETER FunctionUrl
    Public HTTPS URL of your Azure Function webhook endpoint (including the `code` query param).

.PARAMETER TestSuite
    Comma-separated list of test suites to run:
      All, Network, Auth, SharePoint, Function, Webhook
    Defaults to "All".

.PARAMETER AccessToken
    Pre-obtained Bearer access token. When provided, skips authentication tests
    and uses this token for SharePoint tests.

.EXAMPLE
    # Run all tests with service principal
    .\Test-SharePointAuth.ps1 `
        -TenantId  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId  "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -ClientSecret "your-secret" `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite" `
        -SharePointListName "My Webhook List" `
        -FunctionUrl "https://spo-func.azurewebsites.net/api/WebhookTrigger?code=abc123"

.EXAMPLE
    # Run only SharePoint tests using an existing token
    .\Test-SharePointAuth.ps1 `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite" `
        -SharePointListName "My Webhook List" `
        -TestSuite "SharePoint,Webhook" `
        -AccessToken $global:SharePointAccessToken
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)] [string]$TenantId,
    [Parameter(Mandatory = $false)] [string]$ClientId,
    [Parameter(Mandatory = $false)] [string]$ClientSecret,
    [Parameter(Mandatory = $false)] [string]$CertificatePath,
    [Parameter(Mandatory = $false)] [securestring]$CertificatePassword,
    [Parameter(Mandatory = $true)]  [string]$SharePointSiteUrl,
    [Parameter(Mandatory = $false)] [string]$SharePointListName,
    [Parameter(Mandatory = $false)] [string]$FunctionUrl,
    [Parameter(Mandatory = $false)] [string]$TestSuite = "All",
    [Parameter(Mandatory = $false)] [string]$AccessToken
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"  # Tests should not abort on error

#region ── Test framework ────────────────────────────────────────────────────────

$script:TestResults = [System.Collections.Generic.List[PSObject]]::new()
$script:PassCount   = 0
$script:FailCount   = 0
$script:WarnCount   = 0

enum TestStatus { Pass; Fail; Warn; Skip }

function Invoke-Test {
    param(
        [string]$Suite,
        [string]$Name,
        [scriptblock]$Test
    )

    $result = [PSCustomObject]@{
        Suite   = $Suite
        Name    = $Name
        Status  = [TestStatus]::Skip
        Message = ""
        Detail  = ""
    }

    try {
        $output = & $Test
        if ($output -is [hashtable]) {
            $result.Status  = $output.Status
            $result.Message = $output.Message
            $result.Detail  = $output.Detail
        }
        else {
            $result.Status  = [TestStatus]::Pass
            $result.Message = if ($output) { $output.ToString() } else { "OK" }
        }
    }
    catch {
        $result.Status  = [TestStatus]::Fail
        $result.Message = $_.Exception.Message
        $result.Detail  = $_.ScriptStackTrace
    }

    switch ($result.Status) {
        "Pass" { Write-Host "  [PASS] $Name" -ForegroundColor Green;       $script:PassCount++ }
        "Fail" { Write-Host "  [FAIL] $Name : $($result.Message)" -ForegroundColor Red;    $script:FailCount++ }
        "Warn" { Write-Host "  [WARN] $Name : $($result.Message)" -ForegroundColor Yellow; $script:WarnCount++ }
        "Skip" { Write-Host "  [SKIP] $Name" -ForegroundColor DarkGray }
    }

    if ($result.Detail) {
        Write-Host "         $($result.Detail)" -ForegroundColor DarkGray
    }

    $script:TestResults.Add($result)
}

function Write-SuiteHeader {
    param([string]$Name)
    Write-Host ""
    Write-Host "── $Name Tests ──────────────────────────────────────────────────" -ForegroundColor Cyan
}

function Should-RunSuite {
    param([string]$Suite)
    return ($TestSuite -eq "All") -or ($TestSuite -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -eq $Suite })
}

#endregion

#region ── Network Tests ─────────────────────────────────────────────────────────

if (Should-RunSuite "Network") {
    Write-SuiteHeader "Network"

    Invoke-Test -Suite "Network" -Name "HTTPS connectivity to login.microsoftonline.com" -Test {
        $r = Invoke-WebRequest -Uri "https://login.microsoftonline.com" -UseBasicParsing -TimeoutSec 10
        @{ Status = "Pass"; Message = "HTTP $($r.StatusCode)" }
    }

    Invoke-Test -Suite "Network" -Name "HTTPS connectivity to graph.microsoft.com" -Test {
        $r = Invoke-WebRequest -Uri "https://graph.microsoft.com" -UseBasicParsing -TimeoutSec 10
        @{ Status = "Pass"; Message = "HTTP $($r.StatusCode)" }
    }

    $spoHost = ([System.Uri]$SharePointSiteUrl).Host
    Invoke-Test -Suite "Network" -Name "HTTPS connectivity to SharePoint tenant ($spoHost)" -Test {
        $r = Invoke-WebRequest -Uri "https://$spoHost" -UseBasicParsing -TimeoutSec 10
        @{ Status = "Pass"; Message = "HTTP $($r.StatusCode)" }
    }

    if ($FunctionUrl) {
        $funcHost = ([System.Uri]$FunctionUrl).Host
        Invoke-Test -Suite "Network" -Name "HTTPS connectivity to Function App host ($funcHost)" -Test {
            $r = Invoke-WebRequest -Uri "https://$funcHost" -UseBasicParsing -TimeoutSec 10
            @{ Status = "Pass"; Message = "HTTP $($r.StatusCode)" }
        }
    }

    Invoke-Test -Suite "Network" -Name "TLS 1.2 support" -Test {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $r = Invoke-WebRequest -Uri "https://login.microsoftonline.com" -UseBasicParsing -TimeoutSec 10
        @{ Status = "Pass"; Message = "TLS 1.2 working" }
    }
}

#endregion

#region ── Authentication Tests ──────────────────────────────────────────────────

$script:TestToken = $AccessToken

if (Should-RunSuite "Auth") {
    Write-SuiteHeader "Authentication"

    $resource = "$( ([System.Uri]$SharePointSiteUrl).Scheme )://$( ([System.Uri]$SharePointSiteUrl).Host )/"

    # ── Client Secret ──────────────────────────────────────────────────────────
    if ($TenantId -and $ClientId -and $ClientSecret) {
        Invoke-Test -Suite "Auth" -Name "Token via Client Secret" -Test {
            $body = @{
                grant_type    = "client_credentials"
                client_id     = $ClientId
                client_secret = $ClientSecret
                resource      = $resource
            }
            $resp = Invoke-RestMethod `
                -Uri         "https://login.microsoftonline.com/$TenantId/oauth2/token" `
                -Method      Post `
                -Body        $body `
                -ContentType "application/x-www-form-urlencoded" `
                -ErrorAction Stop

            if (-not $resp.access_token) {
                return @{ Status = "Fail"; Message = "Response did not include access_token" }
            }

            $script:TestToken = $resp.access_token
            $expiry = [DateTimeOffset]::UtcNow.AddSeconds([long]$resp.expires_in)
            @{ Status = "Pass"; Message = "Token acquired. Expires at $($expiry.UtcDateTime.ToString('u')) (~$($resp.expires_in)s from now)." }
        }
    }

    # ── Certificate ────────────────────────────────────────────────────────────
    if ($TenantId -and $ClientId -and $CertificatePath) {
        Invoke-Test -Suite "Auth" -Name "Certificate file accessible" -Test {
            if (Test-Path $CertificatePath) {
                @{ Status = "Pass"; Message = "File found: $CertificatePath" }
            }
            else {
                @{ Status = "Fail"; Message = "Certificate file not found: $CertificatePath" }
            }
        }

        Invoke-Test -Suite "Auth" -Name "Certificate validity" -Test {
            $cert = if ($CertificatePassword) {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
                    $CertificatePath, $CertificatePassword)
            }
            else {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CertificatePath)
            }

            $now = Get-Date
            if ($now -lt $cert.NotBefore) {
                return @{ Status = "Fail"; Message = "Certificate is not yet valid (valid from: $($cert.NotBefore))" }
            }
            if ($now -gt $cert.NotAfter) {
                return @{ Status = "Fail"; Message = "Certificate has expired (expired: $($cert.NotAfter))" }
            }
            $daysLeft = ($cert.NotAfter - $now).Days
            $status   = if ($daysLeft -lt 30) { "Warn" } else { "Pass" }
            @{ Status = $status; Message = "Valid. Expires $($cert.NotAfter) ($daysLeft days left). Thumbprint: $($cert.Thumbprint)" }
        }
    }

    # ── Managed Identity ───────────────────────────────────────────────────────
    Invoke-Test -Suite "Auth" -Name "Managed Identity IMDS endpoint reachable" -Test {
        $r = Invoke-WebRequest -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2019-08-01&resource=$([System.Uri]::EscapeDataString($resource))" `
            -Headers @{ Metadata = "true" } -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
        if ($r.StatusCode -eq 200) {
            $token = ($r.Content | ConvertFrom-Json).access_token
            if ($token) {
                if (-not $script:TestToken) { $script:TestToken = $token }
                return @{ Status = "Pass"; Message = "Managed Identity token acquired." }
            }
        }
        @{ Status = "Warn"; Message = "IMDS responded but no token was returned." }
    }

    # ── JWT decode (no signature verification) ────────────────────────────────
    if ($script:TestToken) {
        Invoke-Test -Suite "Auth" -Name "Access token structure" -Test {
            $parts = $script:TestToken.Split(".")
            if ($parts.Count -ne 3) {
                return @{ Status = "Fail"; Message = "Token does not have 3 JWT parts." }
            }
            $pad     = $parts[1].Length % 4
            $padded  = $parts[1] + ("=" * ((4 - $pad) % 4))
            $payload = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($padded)) | ConvertFrom-Json
            $aud     = if ($payload.aud -is [array]) { $payload.aud -join ", " } else { $payload.aud }
            $appId   = $payload.appid ?? $payload.app_id ?? "(n/a)"
            @{ Status = "Pass"; Message = "Audience: $aud | AppId: $appId" }
        }
    }
}

#endregion

#region ── SharePoint Tests ───────────────────────────────────────────────────────

if (Should-RunSuite "SharePoint") {
    Write-SuiteHeader "SharePoint"

    if (-not $script:TestToken) {
        Write-Host "  [SKIP] No access token available – skipping SharePoint tests." -ForegroundColor DarkGray
    }
    else {
        $spHeaders = @{
            Authorization = "Bearer $script:TestToken"
            Accept        = "application/json;odata=nometadata"
        }
        $siteUrlClean = $SharePointSiteUrl.TrimEnd("/")

        Invoke-Test -Suite "SharePoint" -Name "GET /_api/web (site access)" -Test {
            $r = Invoke-RestMethod -Uri "$siteUrlClean/_api/web?`$select=Title,Url,ServerRelativeUrl" `
                -Headers $spHeaders -Method Get -ErrorAction Stop
            @{ Status = "Pass"; Message = "Site: '$($r.Title)' ($($r.Url))" }
        }

        Invoke-Test -Suite "SharePoint" -Name "GET /_api/web/currentuser" -Test {
            $r = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/currentuser?`$select=Title,LoginName,IsSiteAdmin" `
                -Headers $spHeaders -Method Get -ErrorAction Stop
            $adminNote = if ($r.IsSiteAdmin) { " [Site Admin]" } else { "" }
            @{ Status = "Pass"; Message = "Logged in as: $($r.Title) ($($r.LoginName))$adminNote" }
        }

        Invoke-Test -Suite "SharePoint" -Name "GET /_api/web/lists (list permission)" -Test {
            $r = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/lists?`$select=Title,Id&`$top=5" `
                -Headers $spHeaders -Method Get -ErrorAction Stop
            $count = @($r.value).Count
            @{ Status = "Pass"; Message = "Can read lists. Found $count (top 5)." }
        }

        if ($SharePointListName) {
            Invoke-Test -Suite "SharePoint" -Name "GET list '$SharePointListName'" -Test {
                $enc = [System.Uri]::EscapeDataString($SharePointListName)
                $r   = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/lists/getbytitle('$enc')?`$select=Title,Id,ItemCount" `
                    -Headers $spHeaders -Method Get -ErrorAction Stop
                @{ Status = "Pass"; Message = "List ID: $($r.Id), Items: $($r.ItemCount)" }
            }

            Invoke-Test -Suite "SharePoint" -Name "GET webhook subscriptions for '$SharePointListName'" -Test {
                $enc = [System.Uri]::EscapeDataString($SharePointListName)
                $r   = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/lists/getbytitle('$enc')/subscriptions" `
                    -Headers $spHeaders -Method Get -ErrorAction Stop
                $subs = @($r.value)
                if ($subs.Count -eq 0) {
                    return @{ Status = "Warn"; Message = "No webhook subscriptions found on this list." }
                }
                $expiring = $subs | Where-Object { [datetime]$_.expirationDateTime -lt (Get-Date).AddDays(14) }
                if ($expiring.Count -gt 0) {
                    return @{ Status = "Warn"; Message = "$($subs.Count) subscription(s). $($expiring.Count) expire within 14 days!" }
                }
                @{ Status = "Pass"; Message = "$($subs.Count) active subscription(s)." }
            }
        }

        Invoke-Test -Suite "SharePoint" -Name "Write permission check (OPTIONS preflight)" -Test {
            $r = Invoke-WebRequest -Uri "$siteUrlClean/_api/web" -Method Options -UseBasicParsing -Headers $spHeaders -ErrorAction Stop
            @{ Status = "Pass"; Message = "OPTIONS HTTP $($r.StatusCode)" }
        }
    }
}

#endregion

#region ── Azure Function Tests ──────────────────────────────────────────────────

if (Should-RunSuite "Function" -and $FunctionUrl) {
    Write-SuiteHeader "Azure Function"

    Invoke-Test -Suite "Function" -Name "Function URL is HTTPS" -Test {
        if ($FunctionUrl -match "^https://") {
            @{ Status = "Pass"; Message = "URL starts with https://" }
        }
        else {
            @{ Status = "Fail"; Message = "Function URL must use HTTPS. SharePoint will reject plain HTTP." }
        }
    }

    Invoke-Test -Suite "Function" -Name "Validation token handshake (GET)" -Test {
        $testToken = "TestValidationToken_$([System.Guid]::NewGuid().ToString())"
        $url       = if ($FunctionUrl -match "\?") { "$FunctionUrl&validationToken=$testToken" } else { "$FunctionUrl`?validationToken=$testToken" }
        $r         = Invoke-WebRequest -Uri $url -Method Get -UseBasicParsing -ErrorAction Stop
        if ($r.StatusCode -eq 200 -and $r.Content.Trim('"') -eq $testToken) {
            @{ Status = "Pass"; Message = "Validation token echoed correctly (HTTP 200)." }
        }
        elseif ($r.StatusCode -eq 200) {
            @{ Status = "Warn"; Message = "HTTP 200 but response body does not exactly match validation token. Body: $($r.Content.Substring(0,[Math]::Min(100,$r.Content.Length)))" }
        }
        else {
            @{ Status = "Fail"; Message = "Expected HTTP 200, got HTTP $($r.StatusCode)." }
        }
    }

    Invoke-Test -Suite "Function" -Name "Notification POST (mock payload)" -Test {
        $siteId = [System.Guid]::NewGuid().ToString()
        $body   = @{
            value = @(
                @{
                    subscriptionId     = [System.Guid]::NewGuid().ToString()
                    clientState        = "TestState_TroubleshootingScript"
                    expirationDateTime = (Get-Date).AddMonths(3).ToUniversalTime().ToString("o")
                    resource           = "sites/$siteId/lists/$([System.Guid]::NewGuid())"
                    tenantId           = $TenantId
                    siteUrl            = $SharePointSiteUrl
                    webId              = [System.Guid]::NewGuid().ToString()
                    listId             = [System.Guid]::NewGuid().ToString()
                    itemId             = "1"
                    eventType          = "updated"
                }
            )
        } | ConvertTo-Json -Depth 5

        $cleanUrl = $FunctionUrl -replace "\?validationToken=[^&]*", "" -replace "&validationToken=[^&]*", ""
        $r        = Invoke-WebRequest -Uri $cleanUrl -Method Post -Body $body `
            -ContentType "application/json" -UseBasicParsing -ErrorAction Stop

        if ($r.StatusCode -in 200, 202) {
            @{ Status = "Pass"; Message = "Function accepted notification (HTTP $($r.StatusCode))." }
        }
        else {
            @{ Status = "Warn"; Message = "Unexpected HTTP $($r.StatusCode). Body: $($r.Content.Substring(0,[Math]::Min(200,$r.Content.Length)))" }
        }
    }

    Invoke-Test -Suite "Function" -Name "Response time < 5 seconds (SharePoint timeout)" -Test {
        $sw  = [System.Diagnostics.Stopwatch]::StartNew()
        $url = if ($FunctionUrl -match "\?") { "$FunctionUrl&validationToken=timing_test" } else { "$FunctionUrl`?validationToken=timing_test" }
        $r   = Invoke-WebRequest -Uri $url -Method Get -UseBasicParsing -ErrorAction Stop
        $sw.Stop()
        $ms  = $sw.ElapsedMilliseconds

        if ($ms -lt 5000) {
            @{ Status = "Pass"; Message = "Response time: ${ms}ms" }
        }
        else {
            @{ Status = "Warn"; Message = "Response time ${ms}ms exceeds 5s SharePoint timeout. SharePoint may retry or drop webhook." }
        }
    }
}
elseif (Should-RunSuite "Function" -and -not $FunctionUrl) {
    Write-SuiteHeader "Azure Function"
    Write-Host "  [SKIP] FunctionUrl not provided – skipping Function tests." -ForegroundColor DarkGray
}

#endregion

#region ── Webhook Subscription Tests ────────────────────────────────────────────

if (Should-RunSuite "Webhook" -and $SharePointListName -and $script:TestToken) {
    Write-SuiteHeader "Webhook Subscriptions"

    $spHeaders    = @{ Authorization = "Bearer $script:TestToken"; Accept = "application/json;odata=nometadata" }
    $siteUrlClean = $SharePointSiteUrl.TrimEnd("/")
    $enc          = [System.Uri]::EscapeDataString($SharePointListName)

    Invoke-Test -Suite "Webhook" -Name "List active subscriptions" -Test {
        $r    = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/lists/getbytitle('$enc')/subscriptions" `
            -Headers $spHeaders -Method Get -ErrorAction Stop
        $subs = @($r.value)

        if ($subs.Count -eq 0) {
            return @{ Status = "Warn"; Message = "No subscriptions found. Register a webhook first." }
        }

        foreach ($sub in $subs) {
            $exp     = [datetime]$sub.expirationDateTime
            $daysLeft = ($exp - (Get-Date)).Days
            $url     = $sub.notificationUrl
            Write-Host "         ID: $($sub.id)" -ForegroundColor DarkGray
            Write-Host "         URL: $url" -ForegroundColor DarkGray
            Write-Host "         Expires: $exp ($daysLeft days)" -ForegroundColor $(if ($daysLeft -lt 14) { "DarkYellow" } else { "DarkGray" })
        }
        @{ Status = "Pass"; Message = "$($subs.Count) subscription(s) found." }
    }

    Invoke-Test -Suite "Webhook" -Name "No subscriptions expiring within 7 days" -Test {
        $r       = Invoke-RestMethod -Uri "$siteUrlClean/_api/web/lists/getbytitle('$enc')/subscriptions" `
            -Headers $spHeaders -Method Get -ErrorAction Stop
        $subs    = @($r.value)
        $urgentSubs = $subs | Where-Object { ([datetime]$_.expirationDateTime - (Get-Date)).Days -lt 7 }
        if ($urgentSubs.Count -gt 0) {
            return @{ Status = "Warn"; Message = "$($urgentSubs.Count) subscription(s) expire within 7 days. Renew them!" }
        }
        @{ Status = "Pass"; Message = "No subscriptions expiring soon." }
    }
}

#endregion

#region ── Summary ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ""
Write-Host "  Passed  : $($script:PassCount)" -ForegroundColor Green
Write-Host "  Failed  : $($script:FailCount)" -ForegroundColor $(if ($script:FailCount -gt 0) { "Red" } else { "White" })
Write-Host "  Warnings: $($script:WarnCount)" -ForegroundColor $(if ($script:WarnCount -gt 0) { "Yellow" } else { "White" })
Write-Host ""

if ($script:FailCount -gt 0) {
    Write-Host "Failed tests:" -ForegroundColor Red
    $script:TestResults | Where-Object { $_.Status -eq "Fail" } | ForEach-Object {
        Write-Host "  * [$($_.Suite)] $($_.Name): $($_.Message)" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "Troubleshooting tip: See docs/TROUBLESHOOTING.md for detailed guidance." -ForegroundColor DarkYellow
}

if ($script:WarnCount -gt 0) {
    Write-Host "Warnings:" -ForegroundColor Yellow
    $script:TestResults | Where-Object { $_.Status -eq "Warn" } | ForEach-Object {
        Write-Host "  ! [$($_.Suite)] $($_.Name): $($_.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
$overallStatus = if ($script:FailCount -gt 0) { "FAILED" } elseif ($script:WarnCount -gt 0) { "PASSED WITH WARNINGS" } else { "PASSED" }
$color         = if ($script:FailCount -gt 0) { "Red" } elseif ($script:WarnCount -gt 0) { "Yellow" } else { "Green" }
Write-Host "  Overall: $overallStatus" -ForegroundColor $color
Write-Host ""

# Return structured results for pipeline use
return [PSCustomObject]@{
    Pass     = $script:PassCount
    Fail     = $script:FailCount
    Warn     = $script:WarnCount
    Results  = $script:TestResults
    Status   = $overallStatus
}

#endregion
