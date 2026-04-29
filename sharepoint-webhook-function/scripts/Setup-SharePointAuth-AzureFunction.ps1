#Requires -Version 5.1
<#
.SYNOPSIS
    Provisions an Azure Function App with a Managed Identity and grants it
    the necessary SharePoint Online permissions.

.DESCRIPTION
    This script automates the following steps:
      1. Creates (or updates) an Azure Function App with a System-assigned Managed Identity.
      2. Retrieves the service principal object ID for the Managed Identity.
      3. Grants the Managed Identity the required SharePoint API permissions via
         Microsoft Graph (app role assignments).
      4. Optionally sets required application settings on the Function App.
      5. Optionally adds the Managed Identity as a site collection administrator
         on one or more SharePoint site collections.

    Run this script once during initial provisioning. Re-running is idempotent
    (existing role assignments are skipped).

.PARAMETER ResourceGroupName
    Name of the Azure Resource Group that contains the Function App.

.PARAMETER FunctionAppName
    Name of the Azure Function App to configure.

.PARAMETER SubscriptionId
    Azure subscription ID. Defaults to the current Az context subscription.

.PARAMETER SharePointSiteUrl
    One or more SharePoint site URLs to grant the Managed Identity access to.
    Accepts a comma-separated string or an array.

.PARAMETER SharePointPermissionLevel
    SharePoint app role to grant. Valid values:
      - Sites.Selected      (recommended – least privilege)
      - Sites.ReadWrite.All
      - Sites.Manage.All    (required for webhook subscription management)
      - Sites.FullControl.All

.PARAMETER AppSettings
    Hashtable of additional application settings to write to the Function App
    (e.g. SHAREPOINT_SITE_URL, SHAREPOINT_LIST_ID).

.PARAMETER Location
    Azure region for new resources. Defaults to eastus.

.PARAMETER StorageAccountName
    Storage account name for the Function App (created if it does not exist).

.PARAMETER SkipFunctionAppCreation
    When set, assumes the Function App already exists and skips creation.

.EXAMPLE
    # Basic setup – create Function App and grant Sites.Manage.All
    .\Setup-SharePointAuth-AzureFunction.ps1 `
        -ResourceGroupName "SPO-Webhooks-RG" `
        -FunctionAppName   "spo-webhook-func" `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite" `
        -SharePointPermissionLevel "Sites.Manage.All"

.EXAMPLE
    # Existing Function App, write app settings, grant Sites.Selected
    .\Setup-SharePointAuth-AzureFunction.ps1 `
        -ResourceGroupName "SPO-Webhooks-RG" `
        -FunctionAppName   "spo-webhook-func" `
        -SkipFunctionAppCreation `
        -SharePointSiteUrl "https://contoso.sharepoint.com/sites/mysite" `
        -SharePointPermissionLevel "Sites.Selected" `
        -AppSettings @{
            SHAREPOINT_SITE_URL = "https://contoso.sharepoint.com/sites/mysite"
            SHAREPOINT_LIST_ID  = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
        }
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [string]$FunctionAppName,

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $true)]
    [string[]]$SharePointSiteUrl,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Sites.Selected", "Sites.ReadWrite.All", "Sites.Manage.All", "Sites.FullControl.All")]
    [string]$SharePointPermissionLevel = "Sites.Manage.All",

    [Parameter(Mandatory = $false)]
    [hashtable]$AppSettings,

    [Parameter(Mandatory = $false)]
    [string]$Location = "eastus",

    [Parameter(Mandatory = $false)]
    [string]$StorageAccountName,

    [Parameter(Mandatory = $false)]
    [switch]$SkipFunctionAppCreation
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
    param([string]$Step, [string]$Message)
    Write-Host "`nStep $Step : $Message" -ForegroundColor Yellow
}

function Write-Success { param([string]$Message); Write-Host "[+] $Message" -ForegroundColor Green }
function Write-Info    { param([string]$Message); Write-Host "    $Message" -ForegroundColor DarkGray }
function Write-Warn    { param([string]$Message); Write-Host "[!] $Message" -ForegroundColor DarkYellow }

function Assert-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Warn "Module '$Name' is not installed. Installing..."
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -Force
}

#endregion

#region ── Prerequisites ─────────────────────────────────────────────────────────

Write-Banner "Setup SharePoint Auth for Azure Function"

Assert-Module -Name "Az.Accounts"
Assert-Module -Name "Az.Resources"
Assert-Module -Name "Az.Functions"
Assert-Module -Name "Az.Storage"

# Ensure we are logged in
$context = Get-AzContext
if (-not $context) {
    Write-Step "0" "Authenticating to Azure..."
    Connect-AzAccount
    $context = Get-AzContext
}

if ($SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    $context = Get-AzContext
}

Write-Success "Using subscription: $($context.Subscription.Name) ($($context.Subscription.Id))"
Write-Success "Tenant: $($context.Tenant.Id)"

#endregion

#region ── Step 1 – Ensure Resource Group ────────────────────────────────────────

Write-Step "1" "Checking Resource Group '$ResourceGroupName'..."

$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
    if ($PSCmdlet.ShouldProcess($ResourceGroupName, "Create Resource Group")) {
        $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
        Write-Success "Resource group created: $ResourceGroupName ($Location)"
    }
}
else {
    Write-Success "Resource group exists: $ResourceGroupName"
}

#endregion

#region ── Step 2 – Ensure Storage Account ───────────────────────────────────────

if (-not $SkipFunctionAppCreation) {
    Write-Step "2" "Checking Storage Account..."

    if (-not $StorageAccountName) {
        # Generate a storage account name: lowercase letters and digits only, max 24 chars
        $rawName = "spo" + ($FunctionAppName -replace '[^a-z0-9]', '') + "stg"
        $StorageAccountName = $rawName.ToLower().Substring(0, [Math]::Min(24, $rawName.Length))
    }

    $sa = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction SilentlyContinue
    if (-not $sa) {
        if ($PSCmdlet.ShouldProcess($StorageAccountName, "Create Storage Account")) {
            $sa = New-AzStorageAccount `
                -ResourceGroupName $ResourceGroupName `
                -Name              $StorageAccountName `
                -Location          $Location `
                -SkuName           "Standard_LRS" `
                -Kind              "StorageV2"
            Write-Success "Storage account created: $StorageAccountName"
        }
    }
    else {
        Write-Success "Storage account exists: $StorageAccountName"
    }
}

#endregion

#region ── Step 3 – Create / update Function App with Managed Identity ───────────

Write-Step "3" "Configuring Function App '$FunctionAppName'..."

if (-not $SkipFunctionAppCreation) {
    $funcApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ErrorAction SilentlyContinue

    if (-not $funcApp) {
        if ($PSCmdlet.ShouldProcess($FunctionAppName, "Create Function App")) {
            Write-Info "Creating Function App (this may take a minute)..."
            $funcApp = New-AzFunctionApp `
                -ResourceGroupName  $ResourceGroupName `
                -Name               $FunctionAppName `
                -Location           $Location `
                -StorageAccountName $StorageAccountName `
                -Runtime            "dotnet" `
                -RuntimeVersion     "8" `
                -OSType             "Windows" `
                -FunctionsVersion   4
            Write-Success "Function App created: $FunctionAppName"
        }
    }
    else {
        Write-Success "Function App exists: $FunctionAppName"
    }
}
else {
    $funcApp = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName -ErrorAction Stop
    Write-Success "Using existing Function App: $FunctionAppName"
}

# Enable System-assigned Managed Identity
Write-Info "Enabling system-assigned Managed Identity..."
$identityUpdate = Update-AzFunctionApp `
    -ResourceGroupName $ResourceGroupName `
    -Name              $FunctionAppName `
    -IdentityType      SystemAssigned `
    -Force

# Retrieve the principal ID
$funcApp    = Get-AzFunctionApp -ResourceGroupName $ResourceGroupName -Name $FunctionAppName
$principalId = $funcApp.IdentityPrincipalId

if (-not $principalId) {
    throw "Failed to retrieve Managed Identity principal ID for '$FunctionAppName'."
}

Write-Success "Managed Identity enabled. Principal ID: $principalId"

#endregion

#region ── Step 4 – Grant SharePoint permissions via MS Graph ───────────────────

Write-Step "4" "Granting SharePoint permission: $SharePointPermissionLevel"

# Get an access token for Microsoft Graph
$graphToken = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token

$graphHeaders = @{
    Authorization  = "Bearer $graphToken"
    "Content-Type" = "application/json"
}

# Find the SharePoint service principal app roles
Write-Info "Looking up SharePoint service principal..."
$spSearchUrl = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0ff1-ce00-000000000000'&`$select=id,appRoles"
$spSpResult  = Invoke-RestMethod -Uri $spSearchUrl -Headers $graphHeaders -Method Get
$spSp        = $spSpResult.value | Select-Object -First 1

if (-not $spSp) {
    throw "SharePoint service principal not found in this tenant."
}

Write-Info "SharePoint service principal ID: $($spSp.id)"

# Find the matching app role
$appRole = $spSp.appRoles | Where-Object { $_.value -eq $SharePointPermissionLevel }
if (-not $appRole) {
    $available = ($spSp.appRoles | Select-Object -ExpandProperty value) -join ", "
    throw "App role '$SharePointPermissionLevel' not found. Available roles: $available"
}

Write-Info "App role ID for '$SharePointPermissionLevel': $($appRole.id)"

# Check for existing role assignment to avoid duplicates
$existingAssignUrl = "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments"
$existingAssign    = Invoke-RestMethod -Uri $existingAssignUrl -Headers $graphHeaders -Method Get
$alreadyAssigned   = $existingAssign.value | Where-Object { $_.appRoleId -eq $appRole.id -and $_.resourceId -eq $spSp.id }

if ($alreadyAssigned) {
    Write-Warn "Role assignment for '$SharePointPermissionLevel' already exists – skipping."
}
else {
    if ($PSCmdlet.ShouldProcess($FunctionAppName, "Grant app role $SharePointPermissionLevel")) {
        $assignBody = @{
            principalId = $principalId
            resourceId  = $spSp.id
            appRoleId   = $appRole.id
        } | ConvertTo-Json

        $assignUrl = "https://graph.microsoft.com/v1.0/servicePrincipals/$principalId/appRoleAssignments"
        $result    = Invoke-RestMethod -Uri $assignUrl -Method Post -Headers $graphHeaders -Body $assignBody
        Write-Success "Granted '$SharePointPermissionLevel' to Managed Identity (assignment ID: $($result.id))"
    }
}

#endregion

#region ── Step 5 – Grant site collection access (Sites.Selected only) ───────────

if ($SharePointPermissionLevel -eq "Sites.Selected") {
    Write-Step "5" "Granting site-level access (Sites.Selected)..."
    Write-Info  "For Sites.Selected, you must explicitly grant the identity access to each site."
    Write-Info  "This step calls the SharePoint REST API using the current user context."

    foreach ($siteUrl in $SharePointSiteUrl) {
        $siteUrl = $siteUrl.TrimEnd("/")
        Write-Info "Granting 'manage' access to: $siteUrl"

        # Get a SharePoint-scoped token for the current user
        $spoToken = (Get-AzAccessToken -ResourceUrl "$siteUrl").Token

        $permBody = @{
            roles     = @("manage")
            grantedToIdentities = @(
                @{
                    application = @{
                        id          = $principalId
                        displayName = $FunctionAppName
                    }
                }
            )
        } | ConvertTo-Json -Depth 5

        try {
            $permUrl = "$siteUrl/_api/v2.0/sites/root/permissions"
            Invoke-RestMethod -Uri $permUrl -Method Post `
                -Headers @{ Authorization = "Bearer $spoToken"; "Content-Type" = "application/json" } `
                -Body $permBody | Out-Null
            Write-Success "Site access granted: $siteUrl"
        }
        catch {
            Write-Warn "Could not grant site access to $siteUrl – you may need to do this manually."
            Write-Warn "Error: $($_.Exception.Message)"
        }
    }
}
else {
    Write-Info "Skipping per-site permission step (only required for Sites.Selected)."
}

#endregion

#region ── Step 6 – Write Application Settings ───────────────────────────────────

if ($AppSettings -and $AppSettings.Count -gt 0) {
    Write-Step "6" "Writing application settings to Function App..."

    # Merge with existing settings
    $existingSettings = Get-AzFunctionAppSetting -ResourceGroupName $ResourceGroupName -Name $FunctionAppName
    $mergedSettings   = $existingSettings.Clone()

    foreach ($kv in $AppSettings.GetEnumerator()) {
        $mergedSettings[$kv.Key] = $kv.Value
        Write-Info "  $($kv.Key) = $($kv.Value)"
    }

    if ($PSCmdlet.ShouldProcess($FunctionAppName, "Update application settings")) {
        Update-AzFunctionAppSetting `
            -ResourceGroupName $ResourceGroupName `
            -Name              $FunctionAppName `
            -AppSetting        $mergedSettings `
            -Force | Out-Null
        Write-Success "Application settings updated."
    }
}
else {
    Write-Info "No additional application settings provided."
}

#endregion

#region ── Summary ───────────────────────────────────────────────────────────────

Write-Banner "Setup Complete"

Write-Host ""
Write-Host "Function App    : $FunctionAppName" -ForegroundColor White
Write-Host "Resource Group  : $ResourceGroupName" -ForegroundColor White
Write-Host "Principal ID    : $principalId"       -ForegroundColor White
Write-Host "Permission      : $SharePointPermissionLevel" -ForegroundColor White
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "  1. Deploy your function code:  func azure functionapp publish $FunctionAppName"
Write-Host "  2. Get the function URL from Azure Portal or:"
Write-Host "       func azure functionapp list-functions $FunctionAppName"
Write-Host "  3. Register the webhook on your SharePoint list (see docs/WEBHOOK-REGISTRATION.md)"
Write-Host "  4. Test end-to-end with: .\scripts\Test-SharePointAuth.ps1"
Write-Host ""

#endregion
