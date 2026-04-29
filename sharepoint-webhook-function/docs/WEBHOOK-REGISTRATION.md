# Webhook Registration Guide

Step-by-step instructions for registering a SharePoint list webhook that points
to your Azure Function.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Step 1 – Deploy your Azure Function](#step-1--deploy-your-azure-function)
- [Step 2 – Get the function URL](#step-2--get-the-function-url)
- [Step 3 – Get the SharePoint list ID](#step-3--get-the-sharepoint-list-id)
- [Step 4 – Register the webhook subscription](#step-4--register-the-webhook-subscription)
  - [Option A – PnP PowerShell (recommended)](#option-a--pnp-powershell-recommended)
  - [Option B – SharePoint REST API (PowerShell)](#option-b--sharepoint-rest-api-powershell)
  - [Option C – Microsoft Graph REST API](#option-c--microsoft-graph-rest-api)
- [Step 5 – Verify the subscription](#step-5--verify-the-subscription)
- [Step 6 – Test end-to-end](#step-6--test-end-to-end)
- [Renewing a subscription](#renewing-a-subscription)
- [Deleting a subscription](#deleting-a-subscription)
- [Sites.Selected permission](#sites-selected-permission)
- [Automation – register on deploy](#automation--register-on-deploy)

---

## Prerequisites

- Azure Function deployed and publicly reachable over HTTPS.
- Function handles the SharePoint validation token handshake (returns `validationToken` verbatim, HTTP 200, within 5 seconds).
- Service principal (or interactive user) has **SharePoint `Sites.Manage.All`** (or `Sites.Selected` + site access) permission.
- Access token obtained (run `Authenticate-SharePointAppIdentity.ps1` if needed).

---

## Step 1 – Deploy your Azure Function

```bash
# Build and deploy
func azure functionapp publish <function-app-name>
```

Confirm the function is live:
```bash
func azure functionapp list-functions <function-app-name> --output table
```

---

## Step 2 – Get the function URL

### Azure CLI
```bash
az functionapp function show \
  --resource-group <resource-group> \
  --name <function-app-name> \
  --function-name WebhookTrigger \
  --query invokeUrlTemplate \
  --output tsv
```

Append `?code=<function-key>` to the URL. Retrieve the default host key:
```bash
az functionapp keys list \
  --resource-group <resource-group> \
  --name <function-app-name> \
  --query functionKeys.default \
  --output tsv
```

### Azure Portal
1. Open **Function App › Functions › WebhookTrigger**.
2. Click **Get Function URL**.
3. Select **default (function key)** from the dropdown.
4. Copy the full URL.

> **Security note:** The `code` parameter acts as a shared secret. Treat it like a password.

---

## Step 3 – Get the SharePoint list ID

```powershell
$siteUrl  = "https://<tenant>.sharepoint.com/sites/<site>"
$listName = "My Webhook List"
$token    = $global:SharePointAccessToken  # set by Authenticate-SharePointAppIdentity.ps1

$headers = @{
    Authorization = "Bearer $token"
    Accept        = "application/json;odata=nometadata"
}

$r = Invoke-RestMethod `
    -Uri     "$siteUrl/_api/web/lists/getbytitle('$listName')?`$select=Id,Title" `
    -Headers $headers

Write-Host "List ID: $($r.Id)"
```

---

## Step 4 – Register the webhook subscription

SharePoint webhook subscriptions expire after **at most 180 days** and must be renewed.

### Option A – PnP PowerShell (recommended)

```powershell
# Install or update PnP module
Install-Module PnP.PowerShell -Scope CurrentUser -Force

# Connect (interactive)
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive

# Register webhook
$subscription = Add-PnPWebhookSubscription `
    -List            "My Webhook List" `
    -NotificationUrl "https://<func>.azurewebsites.net/api/WebhookTrigger?code=<key>" `
    -ExpirationDays  180 `
    -ClientState     "SecureClientState_ReplaceMe"

Write-Host "Subscription ID: $($subscription.Id)"
Write-Host "Expires        : $($subscription.ExpirationDateTime)"
```

### Option B – SharePoint REST API (PowerShell)

```powershell
$siteUrl  = "https://<tenant>.sharepoint.com/sites/<site>"
$listId   = "<list-id-guid>"
$funcUrl  = "https://<func>.azurewebsites.net/api/WebhookTrigger?code=<key>"
$token    = $global:SharePointAccessToken

$body = @{
    notificationUrl    = $funcUrl
    expirationDateTime = (Get-Date).AddDays(180).ToUniversalTime().ToString("o")
    clientState        = "SecureClientState_ReplaceMe"
} | ConvertTo-Json

$response = Invoke-RestMethod `
    -Uri     "$siteUrl/_api/web/lists('$listId')/subscriptions" `
    -Method  Post `
    -Headers @{
        Authorization  = "Bearer $token"
        "Content-Type" = "application/json"
        Accept         = "application/json;odata=nometadata"
    } `
    -Body $body

Write-Host "Subscription ID: $($response.id)"
```

### Option C – Microsoft Graph REST API

```powershell
$siteHostname = "<tenant>.sharepoint.com"
$siteRelPath  = "/sites/<site>"
$listId       = "<list-id-guid>"
$funcUrl      = "https://<func>.azurewebsites.net/api/WebhookTrigger?code=<key>"
$graphToken   = "<ms-graph-bearer-token>"

$body = @{
    changeType         = "updated"
    notificationUrl    = $funcUrl
    resource           = "https://graph.microsoft.com/v1.0/sites/$siteHostname`:$siteRelPath`:/lists/$listId"
    expirationDateTime = (Get-Date).AddDays(180).ToUniversalTime().ToString("o")
    clientState        = "SecureClientState_ReplaceMe"
} | ConvertTo-Json

$response = Invoke-RestMethod `
    -Uri     "https://graph.microsoft.com/v1.0/subscriptions" `
    -Method  Post `
    -Headers @{
        Authorization  = "Bearer $graphToken"
        "Content-Type" = "application/json"
    } `
    -Body $body

Write-Host "Subscription ID: $($response.id)"
```

> **Note:** Graph subscriptions use a different endpoint (`/subscriptions`) and require `Sites.Manage.All` via Graph, not via the SharePoint API.

---

## Step 5 – Verify the subscription

```powershell
$siteUrl = "https://<tenant>.sharepoint.com/sites/<site>"
$listId  = "<list-id-guid>"
$token   = $global:SharePointAccessToken

$subs = Invoke-RestMethod `
    -Uri     "$siteUrl/_api/web/lists('$listId')/subscriptions" `
    -Headers @{ Authorization = "Bearer $token"; Accept = "application/json;odata=nometadata" }

$subs.value | Format-Table id, notificationUrl, expirationDateTime, clientState
```

Expected output:
```
id                                   notificationUrl                                                   expirationDateTime          clientState
--                                   ---------------                                                   ------------------          -----------
xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx https://<func>.azurewebsites.net/api/WebhookTrigger?code=...     2026-10-26T08:33:31Z        SecureClientState_ReplaceMe
```

---

## Step 6 – Test end-to-end

1. **Create or update an item** in the SharePoint list.
2. **Check Function App logs** (Azure Portal › Function App › Monitor › Logs) for incoming notification.
3. Or run the test script:
   ```powershell
   .\scripts\Test-SharePointAuth.ps1 `
       -SharePointSiteUrl  "https://<tenant>.sharepoint.com/sites/<site>" `
       -SharePointListName "My Webhook List" `
       -FunctionUrl        "https://<func>.azurewebsites.net/api/WebhookTrigger?code=<key>" `
       -TestSuite          "Function,Webhook" `
       -AccessToken        $global:SharePointAccessToken
   ```

---

## Renewing a subscription

Subscriptions expire after a maximum of 180 days. Renew with:

```powershell
# PnP
Set-PnPWebhookSubscription `
    -List           "My Webhook List" `
    -Subscription   "<subscription-id>" `
    -ExpirationDays 180

# REST API
$patchBody = @{
    expirationDateTime = (Get-Date).AddDays(180).ToUniversalTime().ToString("o")
} | ConvertTo-Json

Invoke-RestMethod `
    -Uri    "$siteUrl/_api/web/lists('$listId')/subscriptions('<subscription-id>')" `
    -Method Patch `
    -Headers @{
        Authorization  = "Bearer $token"
        "Content-Type" = "application/json"
        Accept         = "application/json;odata=nometadata"
    } `
    -Body $patchBody
```

**Automate renewal** with a Timer-triggered Azure Function that runs weekly and renews any subscription expiring within 30 days.

---

## Deleting a subscription

```powershell
Invoke-RestMethod `
    -Uri    "$siteUrl/_api/web/lists('$listId')/subscriptions('<subscription-id>')" `
    -Method Delete `
    -Headers @{ Authorization = "Bearer $token" }
```

---

## Sites.Selected permission

When using `Sites.Selected` (least privilege), the app must be explicitly granted access to each site after creating the subscription. This is a two-step process:

### Step 1 – Grant the app permission to the site

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive

Grant-PnPAzureADAppSitePermission `
    -AppId      "<client-id>" `
    -DisplayName "My Webhook App" `
    -Site       "https://<tenant>.sharepoint.com/sites/<site>" `
    -Permissions Manage
```

### Step 2 – Confirm permission

```powershell
Get-PnPAzureADAppSitePermission -Site "https://<tenant>.sharepoint.com/sites/<site>"
```

---

## Automation – register on deploy

Include webhook registration in your CI/CD pipeline (GitHub Actions example):

```yaml
- name: Register SharePoint Webhook
  shell: pwsh
  env:
    TENANT_ID:     ${{ secrets.TENANT_ID }}
    CLIENT_ID:     ${{ secrets.CLIENT_ID }}
    CLIENT_SECRET: ${{ secrets.CLIENT_SECRET }}
    SPO_SITE_URL:  ${{ vars.SPO_SITE_URL }}
    SPO_LIST_ID:   ${{ vars.SPO_LIST_ID }}
    FUNC_URL:      ${{ steps.deploy.outputs.function_url }}
  run: |
    $token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
        -AuthMethod  ServicePrincipal `
        -TenantId    $env:TENANT_ID `
        -ClientId    $env:CLIENT_ID `
        -ClientSecret $env:CLIENT_SECRET `
        -SharePointSiteUrl $env:SPO_SITE_URL

    $body = @{
        notificationUrl    = $env:FUNC_URL
        expirationDateTime = (Get-Date).AddDays(180).ToUniversalTime().ToString("o")
        clientState        = "CI_SecureState"
    } | ConvertTo-Json

    Invoke-RestMethod `
        -Uri    "$env:SPO_SITE_URL/_api/web/lists('$env:SPO_LIST_ID')/subscriptions" `
        -Method Post `
        -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
        -Body $body
```
