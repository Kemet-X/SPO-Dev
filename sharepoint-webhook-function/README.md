# SharePoint Webhook Azure Function

A production-ready reference implementation and setup guide for receiving
SharePoint Online list change notifications via an Azure Function, with complete
authentication support for Service Principal, Managed Identity, and Interactive
login flows.

---

## Table of Contents

- [Overview](#overview)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Part 1 – Azure AD App Registration](#part-1--azure-ad-app-registration)
- [Part 2 – Create the Azure Function Project](#part-2--create-the-azure-function-project)
- [Part 3 – Code Implementation](#part-3--code-implementation)
  - [C# (.NET 8)](#option-a--c-net-8)
  - [TypeScript (Node.js)](#option-b--typescript-nodejs)
- [Part 4 – Configuration](#part-4--configuration)
- [Part 5 – Authentication Methods](#part-5--authentication-methods)
  - [Service Principal – Client Secret](#service-principal--client-secret)
  - [Service Principal – Certificate](#service-principal--certificate)
  - [Managed Identity](#managed-identity)
  - [Interactive / Device Code](#interactive--device-code)
- [Part 6 – Local Testing](#part-6--local-testing)
- [Part 7 – Deploy to Azure](#part-7--deploy-to-azure)
- [Part 8 – Register the Webhook in SharePoint](#part-8--register-the-webhook-in-sharepoint)
- [Part 9 – End-to-End Verification](#part-9--end-to-end-verification)
- [Part 10 – Advanced Scenarios](#part-10--advanced-scenarios)
- [Part 11 – Common Issues & Troubleshooting](#part-11--common-issues--troubleshooting)
- [PowerShell Script Reference](#powershell-script-reference)

---

## Overview

SharePoint Online allows you to subscribe to change notifications (webhooks) on
list items. When an item is added, updated, or deleted SharePoint sends an HTTP
POST notification to a URL you provide – in this case an Azure Function endpoint.

```
SharePoint List change
        │
        ▼
SharePoint Webhook Service
        │  POST notification (JSON)
        ▼
Azure Function (WebhookTrigger)
        │
        ├─► Return HTTP 200 immediately (< 5 seconds)
        │
        └─► Enqueue to Azure Queue Storage (async processing)
                │
                ▼
        QueueTrigger Function
                │
                └─► Business logic (Graph API, email, DB, etc.)
```

---

## Project Structure

```
sharepoint-webhook-function/
├── scripts/
│   ├── Authenticate-SharePointAppIdentity.ps1   # Authentication helper
│   ├── Setup-SharePointAuth-AzureFunction.ps1   # Provision Function App + permissions
│   └── Test-SharePointAuth.ps1                  # Full test & troubleshooting suite
├── docs/
│   ├── TROUBLESHOOTING.md                       # Detailed issue resolution guide
│   └── WEBHOOK-REGISTRATION.md                  # Webhook subscription guide
├── config-sample.json                           # Configuration template
├── .env.example                                 # Environment variables template
└── README.md                                    # This file
```

---

## Prerequisites

| Requirement | Version / Notes |
|---|---|
| Azure subscription | Free tier works for development |
| Visual Studio Code | Latest stable |
| Azure Functions Core Tools | v4 (`npm i -g azure-functions-core-tools@4`) |
| .NET SDK | 8.0+ (for C# implementation) |
| Node.js | 18 LTS+ (for TypeScript implementation) |
| Azure CLI | Latest (`az --version`) |
| PowerShell | 5.1+ (7+ recommended) |
| Az PowerShell module | `Install-Module Az -Scope CurrentUser` |
| Azurite (local storage emulator) | `npm i -g azurite` |

---

## Part 1 – Azure AD App Registration

### 1.1 Create the registration

```powershell
# Log in
az login

# Create app registration
$app = az ad app create `
    --display-name "SharePoint-Webhook-Function" `
    --query "{appId:appId, objectId:id}" `
    --output json | ConvertFrom-Json

Write-Host "Client ID : $($app.appId)"
Write-Host "Object ID : $($app.objectId)"
```

### 1.2 Create a service principal

```powershell
az ad sp create --id $app.appId
```

### 1.3 Grant SharePoint permissions

1. Open **Azure Portal › Azure Active Directory › App registrations**.
2. Select your app → **API permissions → Add a permission**.
3. Choose **SharePoint** (not Microsoft Graph).
4. Select **Application permissions → Sites.Manage.All**.
5. Click **Grant admin consent**.

> **Least privilege:** For production, prefer `Sites.Selected` – it restricts the
> app to only the sites you explicitly grant. See
> [docs/WEBHOOK-REGISTRATION.md](docs/WEBHOOK-REGISTRATION.md#sites-selected-permission).

### 1.4 Create a client secret (Service Principal method)

```powershell
$secret = az ad app credential reset `
    --id $app.appId `
    --years 1 `
    --query password `
    --output tsv

Write-Host "Client Secret: $secret"
# ⚠ Store this securely – it is shown only once
```

### 1.5 Create and upload a certificate (alternative to secret)

```powershell
# Generate self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject       "CN=SharePoint-Webhook-Function" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export PFX (with private key)
$pwd = Read-Host -Prompt "PFX password" -AsSecureString
Export-PfxCertificate -Cert $cert -FilePath ".\spwebhook.pfx" -Password $pwd

# Upload public key to app registration
.\Scripts\Configure-AppRegistration.ps1 `
    -CertificatePath ".\spwebhook.pfx" `
    -ApplicationId   $app.appId
```

---

## Part 2 – Create the Azure Function Project

```bash
# C# (.NET 8) – isolated worker model
func init SharePointWebhookFunction --worker-runtime dotnet-isolated --target-framework net8.0
cd SharePointWebhookFunction
func new --name WebhookTrigger --template "HTTP trigger" --authlevel function

# TypeScript (Node.js 18)
func init SharePointWebhookFunction --worker-runtime node --language typescript
cd SharePointWebhookFunction
func new --name WebhookTrigger --template "HTTP trigger" --authlevel function
```

---

## Part 3 – Code Implementation

### Option A – C# (.NET 8)

```csharp
// WebhookTrigger.cs
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System.Net;
using System.Text.Json;

public class WebhookTrigger
{
    private readonly ILogger<WebhookTrigger> _logger;

    public WebhookTrigger(ILogger<WebhookTrigger> logger) => _logger = logger;

    [Function("WebhookTrigger")]
    public async Task<HttpResponseData> RunAsync(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
        _logger.LogInformation("SharePoint webhook triggered");

        // ── Validation handshake ────────────────────────────────────────────
        // SharePoint sends GET ?validationToken=<token> when registering
        var validationToken = req.Query["validationToken"];
        if (!string.IsNullOrEmpty(validationToken))
        {
            _logger.LogInformation("Validation handshake received");
            var validationResponse = req.CreateResponse(HttpStatusCode.OK);
            validationResponse.Headers.Add("Content-Type", "text/plain");
            await validationResponse.WriteStringAsync(validationToken);
            return validationResponse;
        }

        // ── Notification payload ────────────────────────────────────────────
        string body = await new StreamReader(req.Body).ReadToEndAsync();
        _logger.LogInformation("Notification received: {Body}", body);

        try
        {
            var notification = JsonSerializer.Deserialize<SharePointNotificationEnvelope>(body,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (notification?.Value != null)
            {
                foreach (var item in notification.Value)
                {
                    _logger.LogInformation(
                        "List {ListId} changed on site {SiteUrl} (subscription {SubscriptionId})",
                        item.ListId, item.SiteUrl, item.SubscriptionId);

                    // TODO: enqueue for async processing or call business logic
                }
            }
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Failed to deserialise notification body");
        }

        // ⚠ Must return HTTP 200 within 5 seconds or SharePoint will retry
        return req.CreateResponse(HttpStatusCode.OK);
    }
}

// ── Models ──────────────────────────────────────────────────────────────────

public record SharePointNotificationEnvelope(
    [property: System.Text.Json.Serialization.JsonPropertyName("value")]
    List<SharePointNotificationItem> Value);

public record SharePointNotificationItem(
    string SubscriptionId,
    string ClientState,
    DateTime ExpirationDateTime,
    string Resource,
    string TenantId,
    string SiteUrl,
    string WebId,
    string ListId,
    string? ItemId,
    string? EventType);
```

Add required NuGet packages to your `.csproj`:

```xml
<ItemGroup>
  <PackageReference Include="Microsoft.Azure.Functions.Worker"           Version="1.23.0" />
  <PackageReference Include="Microsoft.Azure.Functions.Worker.Extensions.Http" Version="3.2.0" />
  <PackageReference Include="Microsoft.Azure.Functions.Worker.Sdk"       Version="1.17.4" />
  <PackageReference Include="Azure.Storage.Queues"                       Version="12.21.0" />
</ItemGroup>
```

---

### Option B – TypeScript (Node.js)

```typescript
// src/functions/WebhookTrigger.ts
import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

interface SharePointNotificationItem {
    subscriptionId: string;
    clientState: string;
    expirationDateTime: string;
    resource: string;
    tenantId: string;
    siteUrl: string;
    webId: string;
    listId: string;
    itemId?: string;
    eventType?: string;
}

async function webhookTrigger(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {

    context.log("SharePoint webhook triggered");

    // ── Validation handshake ────────────────────────────────────────────────
    const validationToken = request.query.get("validationToken");
    if (validationToken) {
        context.log("Validation handshake received");
        return {
            status: 200,
            headers: { "Content-Type": "text/plain" },
            body: validationToken
        };
    }

    // ── Notification payload ────────────────────────────────────────────────
    try {
        const payload = await request.json() as { value: SharePointNotificationItem[] };

        for (const item of payload?.value ?? []) {
            context.log(`List ${item.listId} changed on ${item.siteUrl}`);
            // TODO: enqueue or call business logic
        }
    } catch (err) {
        context.error("Failed to parse notification", err);
    }

    // ⚠ Must return 200 within 5 seconds
    return { status: 200 };
}

app.http("WebhookTrigger", {
    methods: ["GET", "POST"],
    authLevel: "function",
    handler: webhookTrigger
});
```

`package.json` dependencies:

```json
{
  "dependencies": {
    "@azure/functions": "^4.5.0",
    "@azure/storage-queue": "^12.21.0"
  },
  "devDependencies": {
    "@types/node": "^20.0.0",
    "typescript": "^5.4.0"
  }
}
```

---

## Part 4 – Configuration

### local.settings.json (not committed)

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "TENANT_ID": "<tenant-id>",
    "CLIENT_ID": "<client-id>",
    "CLIENT_SECRET": "<client-secret>",
    "SHAREPOINT_SITE_URL": "https://<tenant>.sharepoint.com/sites/<site>",
    "SHAREPOINT_LIST_ID": "<list-id-guid>",
    "WEBHOOK_CLIENT_STATE": "<random-secure-string>"
  }
}
```

Copy [`.env.example`](.env.example) → `.env` and fill in values for local PowerShell scripts.

---

## Part 5 – Authentication Methods

All three methods are supported by `scripts/Authenticate-SharePointAppIdentity.ps1`.

### Service Principal – Client Secret

```powershell
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod    ServicePrincipal `
    -TenantId      $env:TENANT_ID `
    -ClientId      $env:CLIENT_ID `
    -ClientSecret  $env:CLIENT_SECRET `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>"
```

Best for: CI/CD pipelines, automated scripts.

### Service Principal – Certificate

More secure than client secrets (private key never transmitted):

```powershell
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod          ServicePrincipal `
    -TenantId            $env:TENANT_ID `
    -ClientId            $env:CLIENT_ID `
    -CertificatePath     ".\spwebhook.pfx" `
    -CertificatePassword (Read-Host -AsSecureString "Cert password") `
    -SharePointSiteUrl   "https://<tenant>.sharepoint.com/sites/<site>"
```

Or use a certificate already installed in the Windows certificate store:

```powershell
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod            ServicePrincipal `
    -TenantId              $env:TENANT_ID `
    -ClientId              $env:CLIENT_ID `
    -CertificateThumbprint "AABBCCDDEEFF..." `
    -SharePointSiteUrl     "https://<tenant>.sharepoint.com/sites/<site>"
```

### Managed Identity

No credentials required – the Azure resource's identity is used automatically.

```powershell
# Run this from within an Azure VM, App Service, or Function App
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod        ManagedIdentity `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>"

# User-assigned identity (provide the client ID of the identity)
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod        ManagedIdentity `
    -ClientId          "<user-assigned-identity-client-id>" `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>"
```

Provision the Managed Identity and grant SharePoint permissions in one step:

```powershell
.\scripts\Setup-SharePointAuth-AzureFunction.ps1 `
    -ResourceGroupName         "SPO-Webhooks-RG" `
    -FunctionAppName           "spo-webhook-func" `
    -SharePointSiteUrl         "https://<tenant>.sharepoint.com/sites/<site>" `
    -SharePointPermissionLevel "Sites.Manage.All" `
    -AppSettings @{
        SHAREPOINT_SITE_URL = "https://<tenant>.sharepoint.com/sites/<site>"
        SHAREPOINT_LIST_ID  = "<list-id>"
    }
```

### Interactive / Device Code

Useful for admin tasks or one-off operations:

```powershell
$token = .\scripts\Authenticate-SharePointAppIdentity.ps1 `
    -AuthMethod    Interactive `
    -TenantId      $env:TENANT_ID `
    -ClientId      $env:CLIENT_ID `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>"
# Follow the printed device code instructions in a browser
```

---

## Part 6 – Local Testing

### 6.1 Start local dependencies

```bash
# Start Azurite (local Azure Storage emulator)
azurite --silent --location ./azurite-data &

# Start the function
func start
```

### 6.2 Test the validation handshake

```bash
# Should echo the token back as plain text
curl "http://localhost:7071/api/WebhookTrigger?code=<key>&validationToken=test123"
# Expected: test123
```

Using PowerShell:

```powershell
$r = Invoke-WebRequest -Uri "http://localhost:7071/api/WebhookTrigger?validationToken=test123"
$r.Content  # Should be: test123
```

### 6.3 Test a notification payload

```powershell
$body = @{
    value = @(
        @{
            subscriptionId     = [guid]::NewGuid().ToString()
            clientState        = "SecureClientState_ReplaceMe"
            expirationDateTime = (Get-Date).AddMonths(3).ToString("o")
            resource           = "sites/00000000/lists/11111111"
            tenantId           = "your-tenant-id"
            siteUrl            = "https://<tenant>.sharepoint.com/sites/<site>"
            webId              = [guid]::NewGuid().ToString()
            listId             = [guid]::NewGuid().ToString()
            itemId             = "42"
            eventType          = "updated"
        }
    )
} | ConvertTo-Json -Depth 5

Invoke-RestMethod `
    -Uri         "http://localhost:7071/api/WebhookTrigger" `
    -Method      Post `
    -Body        $body `
    -ContentType "application/json"
# Expected: HTTP 200
```

### 6.4 Expose local function to the internet (for subscription registration)

```bash
# Using ngrok
ngrok http 7071
# Copy the https://... URL and use it as your notificationUrl

# Using Azure Dev Tunnels
devtunnel host -p 7071 --allow-anonymous
```

### 6.5 Run the full diagnostic test suite

```powershell
.\scripts\Test-SharePointAuth.ps1 `
    -TenantId          $env:TENANT_ID `
    -ClientId          $env:CLIENT_ID `
    -ClientSecret      $env:CLIENT_SECRET `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>" `
    -SharePointListName "My Webhook List" `
    -FunctionUrl       "http://localhost:7071/api/WebhookTrigger" `
    -TestSuite         "Network,Auth,SharePoint,Function"
```

---

## Part 7 – Deploy to Azure

### 7.1 One-command setup

```powershell
.\scripts\Setup-SharePointAuth-AzureFunction.ps1 `
    -ResourceGroupName "SPO-Webhooks-RG" `
    -FunctionAppName   "spo-webhook-func" `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>" `
    -SharePointPermissionLevel "Sites.Manage.All"
```

### 7.2 Manual setup via Azure CLI

```bash
# 1 – Resource group
az group create --name SPO-Webhooks-RG --location eastus

# 2 – Storage account
az storage account create \
    --name spowebhookfuncstg \
    --resource-group SPO-Webhooks-RG \
    --location eastus \
    --sku Standard_LRS

# 3 – Function App (.NET 8)
az functionapp create \
    --resource-group SPO-Webhooks-RG \
    --consumption-plan-location eastus \
    --runtime dotnet-isolated \
    --runtime-version 8 \
    --functions-version 4 \
    --name spo-webhook-func \
    --storage-account spowebhookfuncstg \
    --assign-identity [system]

# 4 – Set application settings
az functionapp config appsettings set \
    --name spo-webhook-func \
    --resource-group SPO-Webhooks-RG \
    --settings \
        TENANT_ID="<tenant-id>" \
        CLIENT_ID="<client-id>" \
        SHAREPOINT_SITE_URL="https://<tenant>.sharepoint.com/sites/<site>" \
        SHAREPOINT_LIST_ID="<list-id>" \
        WEBHOOK_CLIENT_STATE="<random-secure-string>"

# 5 – Deploy function code
func azure functionapp publish spo-webhook-func
```

### 7.3 Retrieve the function URL

```bash
func azure functionapp list-functions spo-webhook-func --output table

# Or get the URL directly:
az functionapp function show \
    --resource-group SPO-Webhooks-RG \
    --name spo-webhook-func \
    --function-name WebhookTrigger \
    --query invokeUrlTemplate --output tsv
```

---

## Part 8 – Register the Webhook in SharePoint

See the full guide: [docs/WEBHOOK-REGISTRATION.md](docs/WEBHOOK-REGISTRATION.md)

### Quick start – PnP PowerShell

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive

Add-PnPWebhookSubscription `
    -List            "My Webhook List" `
    -NotificationUrl "https://spo-webhook-func.azurewebsites.net/api/WebhookTrigger?code=<key>" `
    -ExpirationDays  180 `
    -ClientState     "SecureClientState_ReplaceMe"
```

### Quick start – REST API

```powershell
# Assumes $token set by Authenticate-SharePointAppIdentity.ps1

$body = @{
    notificationUrl    = "https://spo-webhook-func.azurewebsites.net/api/WebhookTrigger?code=<key>"
    expirationDateTime = (Get-Date).AddDays(180).ToUniversalTime().ToString("o")
    clientState        = "SecureClientState_ReplaceMe"
} | ConvertTo-Json

Invoke-RestMethod `
    -Uri     "https://<tenant>.sharepoint.com/sites/<site>/_api/web/lists('<list-id>')/subscriptions" `
    -Method  Post `
    -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
    -Body    $body
```

---

## Part 9 – End-to-End Verification

1. **Create or modify a list item** in SharePoint.
2. **Check function logs** in Azure Portal → Function App → Monitor → Logs.
3. **Run the test script:**

```powershell
.\scripts\Test-SharePointAuth.ps1 `
    -TenantId          $env:TENANT_ID `
    -ClientId          $env:CLIENT_ID `
    -ClientSecret      $env:CLIENT_SECRET `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>" `
    -SharePointListName "My Webhook List" `
    -FunctionUrl       "https://spo-webhook-func.azurewebsites.net/api/WebhookTrigger?code=<key>" `
    -TestSuite         "All"
```

---

## Part 10 – Advanced Scenarios

### Async processing with Azure Queue Storage

Return HTTP 200 immediately and process the notification in a separate function:

```csharp
// WebhookTrigger.cs – enqueue for async processing
[Function("WebhookTrigger")]
[QueueOutput("sharepoint-notifications", Connection = "AzureWebJobsStorage")]
public async Task<QueueOutputType> RunAsync(
    [HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
{
    if (req.Query["validationToken"] is { } vt)
    {
        var vr = req.CreateResponse(HttpStatusCode.OK);
        vr.Headers.Add("Content-Type", "text/plain");
        await vr.WriteStringAsync(vt);
        return new QueueOutputType { HttpResponse = vr };
    }

    var body = await new StreamReader(req.Body).ReadToEndAsync();
    var ok   = req.CreateResponse(HttpStatusCode.OK);
    return new QueueOutputType { HttpResponse = ok, QueueMessage = body };
}

// ProcessNotification.cs – long-running processing
[Function("ProcessNotification")]
public async Task RunAsync(
    [QueueTrigger("sharepoint-notifications")] string message,
    FunctionContext context)
{
    var logger = context.GetLogger<ProcessNotification>();
    // Parse and process the notification
    logger.LogInformation("Processing: {Message}", message);
}

public record QueueOutputType
{
    [QueueOutput("sharepoint-notifications")]
    public string? QueueMessage { get; init; }
    public HttpResponseData HttpResponse { get; init; } = null!;
}
```

### Validate clientState to prevent spoofed notifications

```csharp
var expectedClientState = Environment.GetEnvironmentVariable("WEBHOOK_CLIENT_STATE");
if (item.ClientState != expectedClientState)
{
    _logger.LogWarning("Rejected notification with unexpected clientState");
    continue;
}
```

### Subscription renewal timer

```csharp
[Function("RenewWebhookSubscriptions")]
public async Task RunAsync([TimerTrigger("0 0 9 * * 1")] TimerInfo timer) // Every Monday at 09:00
{
    // List subscriptions, renew any expiring within 30 days
}
```

---

## Part 11 – Common Issues & Troubleshooting

| Issue | Cause | Solution |
|-------|-------|---------|
| `401 Unauthorized` from SharePoint | Wrong token audience or missing permission | Token resource must be `https://<tenant>.sharepoint.com/`, grant `Sites.Manage.All` |
| `Subscription validation request failed` | Function not returning `validationToken` correctly | Return token as plain text, HTTP 200, within 5 seconds |
| Webhook not firing | Subscription expired or function URL unreachable | Check subscription, renew if expired, verify function is reachable |
| `notificationUrl is not valid` | URL not HTTPS or validation handshake fails | Use HTTPS only; test with ngrok locally |
| Timeout on cold start | Consumption plan cold start > 5 s | Use Premium plan or Always On |
| `AADSTS700016` | Wrong ClientId or TenantId | Verify app registration in Azure Portal |
| `AADSTS65001` | Admin consent not granted | Grant admin consent in API permissions |
| Certificate private key error | Wrong storage flags | Use `MachineKeySet \| PersistKeySet` flags |
| `403 Forbidden` on subscriptions endpoint | `Sites.ReadWrite.All` is insufficient | Upgrade to `Sites.Manage.All` |

For detailed resolution steps see [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

---

## PowerShell Script Reference

### `Authenticate-SharePointAppIdentity.ps1`

Acquires a SharePoint Bearer token using the specified authentication method.

```
PARAMETERS
  -AuthMethod         ServicePrincipal | ManagedIdentity | Interactive  [Required]
  -TenantId           Azure AD tenant ID                                 [Required for SP/Interactive]
  -ClientId           App registration client ID                         [Required for SP/Interactive]
  -ClientSecret       Client secret                                      [SP – secret method]
  -CertificatePath    Path to PFX file                                   [SP – cert method]
  -CertificatePassword  SecureString password for PFX                   [SP – cert method]
  -CertificateThumbprint  Cert thumbprint from Windows store             [SP – cert method]
  -SharePointSiteUrl  Target SharePoint site URL                         [Required]
  -OutputTokenToFile  File path to write the token                       [Optional]

OUTPUT
  Returns the access token string and sets $global:SharePointAccessToken
```

### `Setup-SharePointAuth-AzureFunction.ps1`

Creates or configures an Azure Function App with Managed Identity and grants SharePoint permissions.

```
PARAMETERS
  -ResourceGroupName          [Required]
  -FunctionAppName            [Required]
  -SharePointSiteUrl          [Required]  (single or array)
  -SharePointPermissionLevel  Sites.Selected | Sites.ReadWrite.All | Sites.Manage.All | Sites.FullControl.All
  -AppSettings                Hashtable of Function App settings to write
  -Location                   Azure region (default: eastus)
  -StorageAccountName         Storage account (auto-generated if omitted)
  -SkipFunctionAppCreation    Switch – skip creation, configure existing app
  -SubscriptionId             Azure subscription (defaults to current context)
```

### `Test-SharePointAuth.ps1`

Runs automated connectivity, authentication, SharePoint API, function endpoint, and webhook subscription checks.

```
PARAMETERS
  -TenantId, -ClientId, -ClientSecret  For Auth tests
  -CertificatePath, -CertificatePassword  For Auth tests (cert)
  -SharePointSiteUrl   [Required]
  -SharePointListName  For list-level tests
  -FunctionUrl         For Function and Webhook tests
  -TestSuite           All | Network | Auth | SharePoint | Function | Webhook
  -AccessToken         Use a pre-obtained token (skip Auth tests)

OUTPUT
  PSCustomObject { Pass; Fail; Warn; Results; Status }
```

---

## Related Resources

- [SharePoint webhooks overview](https://learn.microsoft.com/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks)
- [Azure Functions documentation](https://learn.microsoft.com/azure/azure-functions/)
- [PnP PowerShell reference](https://pnp.github.io/powershell/)
- [Microsoft identity platform – OAuth 2.0 client credentials](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
- [Azure Managed Identities](https://learn.microsoft.com/entra/identity/managed-identities-azure-resources/overview)
