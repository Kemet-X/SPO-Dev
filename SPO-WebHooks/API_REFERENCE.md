# API Reference — SPO-WebHooks

Complete reference documentation for the SPO-WebHooks library and REST API.

---

## Table of Contents

1. [Core Library](#1-core-library)
   - [AzureAdAuthenticationService](#azureadauthenticationservice)
   - [SharePointWebhookService](#sharepointwebhookservice)
   - [WebhookNotificationHandler](#webhooknotificationhandler)
2. [Models](#2-models)
   - [WebhookSubscription](#webhooksubscription)
   - [WebhookNotification / NotificationData](#webhooknotification--notificationdata)
3. [REST API Endpoints](#3-rest-api-endpoints)
   - [POST /api/webhook](#post-apiwebhook)
   - [GET /api/webhook/health](#get-apiwebhookhealth)
4. [Configuration Reference](#4-configuration-reference)
5. [Error Codes](#5-error-codes)

---

## 1. Core Library

### AzureAdAuthenticationService

**Namespace:** `SPO.Webhooks.Core.Services`

Provides Azure AD certificate-based (OAuth 2.0 client credentials) authentication.

#### Constructor

```csharp
public AzureAdAuthenticationService(
    string tenantId,
    string clientId,
    X509Certificate2 certificate,
    ILogger<AzureAdAuthenticationService> logger)
```

| Parameter | Type | Description |
|---|---|---|
| `tenantId` | `string` | Azure AD Directory (tenant) ID |
| `clientId` | `string` | Azure AD Application (client) ID |
| `certificate` | `X509Certificate2` | Certificate with private key for client assertion |
| `logger` | `ILogger<…>` | Logger instance (use dependency injection) |

**Throws:** `ArgumentException` if `tenantId` or `clientId` is null/empty/whitespace.  
**Throws:** `ArgumentNullException` if `certificate` is null.

---

#### Methods

##### `GetAccessTokenAsync()`

Acquires an access token scoped to Microsoft Graph (`.default`).

```csharp
public Task<string> GetAccessTokenAsync()
```

**Returns:** Bearer access token string.  
**Throws:** `MsalServiceException` on Azure AD error. `MsalClientException` on local MSAL error.

---

##### `GetSharePointAccessTokenAsync(string sharePointTenantUrl)`

Acquires an access token scoped to a specific SharePoint tenant.

```csharp
public Task<string> GetSharePointAccessTokenAsync(string sharePointTenantUrl)
```

| Parameter | Type | Description |
|---|---|---|
| `sharePointTenantUrl` | `string` | Root URL, e.g. `https://contoso.sharepoint.com` |

**Returns:** Bearer access token string scoped to SharePoint.

---

##### `GetAuthorityUrl()`

Returns the OAuth 2.0 token endpoint for the configured tenant.

```csharp
public string GetAuthorityUrl()
// Returns: "https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token"
```

---

### SharePointWebhookService

**Namespace:** `SPO.Webhooks.Core.Services`

Manages SharePoint Online webhook subscriptions using the SharePoint REST API.

#### Constructor

```csharp
public SharePointWebhookService(
    HttpClient httpClient,
    string sharePointSiteUrl,
    string listName,
    ILogger<SharePointWebhookService> logger)
```

| Parameter | Type | Description |
|---|---|---|
| `httpClient` | `HttpClient` | HTTP client (inject via `IHttpClientFactory`) |
| `sharePointSiteUrl` | `string` | Full site URL, e.g. `https://contoso.sharepoint.com/sites/mysite` |
| `listName` | `string` | Display name of the target list |
| `logger` | `ILogger<…>` | Logger instance |

---

#### Methods

##### `SetAccessToken(string accessToken)`

Sets the Bearer token used for all subsequent API calls.

```csharp
public void SetAccessToken(string accessToken)
```

---

##### `CreateSubscriptionAsync(...)`

Creates a new webhook subscription.

```csharp
public Task<WebhookSubscription> CreateSubscriptionAsync(
    string notificationUrl,
    string clientState = "",
    int expirationInDays = 180)
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `notificationUrl` | `string` | required | Public HTTPS URL of your webhook receiver |
| `clientState` | `string` | `""` | Secret for validating incoming notifications |
| `expirationInDays` | `int` | `180` | Days until expiry (1–180) |

**Returns:** Created `WebhookSubscription`.  
**Throws:** `ArgumentOutOfRangeException` if `expirationInDays` is outside 1–180.

---

##### `GetSubscriptionsAsync()`

Retrieves all webhook subscriptions for the configured list.

```csharp
public Task<List<WebhookSubscription>> GetSubscriptionsAsync()
```

**Returns:** List of `WebhookSubscription` objects.

---

##### `RenewSubscriptionAsync(string subscriptionId, int expirationInDays = 180)`

Updates the expiration of an existing subscription.

```csharp
public Task<WebhookSubscription> RenewSubscriptionAsync(
    string subscriptionId,
    int expirationInDays = 180)
```

**Returns:** Updated `WebhookSubscription`.

---

##### `DeleteSubscriptionAsync(string subscriptionId)`

Permanently removes a webhook subscription.

```csharp
public Task DeleteSubscriptionAsync(string subscriptionId)
```

---

### WebhookNotificationHandler

**Namespace:** `SPO.Webhooks.Core.Services`

Validates and processes incoming SharePoint webhook notification payloads.

#### Constructor

```csharp
public WebhookNotificationHandler(
    string clientState,
    ILogger<WebhookNotificationHandler> logger)
```

| Parameter | Type | Description |
|---|---|---|
| `clientState` | `string` | Secret used to verify HMAC-SHA256 signatures. Pass empty string to skip signature validation. |
| `logger` | `ILogger<…>` | Logger instance |

---

#### Methods

##### `ValidateNotification(string payload, string? signatureHeader = null)`

Validates a raw notification payload.

```csharp
public bool ValidateNotification(string payload, string? signatureHeader = null)
```

| Parameter | Description |
|---|---|
| `payload` | Raw JSON body of the incoming POST |
| `signatureHeader` | Value of `X-SP-Webhook-Signature` header (optional) |

**Returns:** `true` if the notification passes all validation checks.

**Validation checks:**
1. Payload is not null/empty.
2. HMAC-SHA256 signature matches (when both `clientState` and `signatureHeader` are provided).
3. Each `clientState` field in the payload matches the configured secret.

---

##### `ParseNotification(string payload)`

Deserializes a raw notification payload.

```csharp
public WebhookNotification ParseNotification(string payload)
```

**Returns:** Parsed `WebhookNotification` object.  
**Throws:** `JsonException` on malformed JSON. `InvalidOperationException` if result is null.

---

##### `ProcessNotificationsAsync(WebhookNotification, Func<NotificationData, Task>)`

Invokes the provided handler for each notification entry. Continues processing remaining entries even if one handler throws.

```csharp
public Task ProcessNotificationsAsync(
    WebhookNotification notification,
    Func<NotificationData, Task> handler)
```

---

## 2. Models

### WebhookSubscription

**Namespace:** `SPO.Webhooks.Core.Models`

```csharp
public class WebhookSubscription
{
    public string Id { get; set; }
    public string Resource { get; set; }
    public string NotificationUrl { get; set; }
    public string ClientState { get; set; }
    public DateTime ExpirationDateTime { get; set; }
    public DateTime CreationTime { get; set; }
}
```

| Property | JSON key | Description |
|---|---|---|
| `Id` | `id` | GUID identifying the subscription |
| `Resource` | `resource` | SharePoint list REST API URL |
| `NotificationUrl` | `notificationUrl` | Receiver endpoint URL |
| `ClientState` | `clientState` | Secret for notification validation |
| `ExpirationDateTime` | `expirationDateTime` | UTC expiry timestamp |
| `CreationTime` | `creationTime` | UTC creation timestamp |

---

### WebhookNotification / NotificationData

**Namespace:** `SPO.Webhooks.Core.Models`

```csharp
public class WebhookNotification
{
    public List<NotificationData> Value { get; set; }
}

public class NotificationData
{
    public string SubscriptionId { get; set; }
    public string ClientState { get; set; }
    public DateTime ExpirationDateTime { get; set; }
    public string Resource { get; set; }
    public string TenantId { get; set; }
    public string SiteUrl { get; set; }
    public string WebId { get; set; }
}
```

---

## 3. REST API Endpoints

Base URL: `https://<your-domain>/api`

---

### POST /api/webhook

Receives SharePoint webhook validation challenges and notifications.

#### Validation Handshake (GET-like POST with query param)

SharePoint sends this when creating a subscription:

```
POST /api/webhook?validationtoken=<token>
```

**Response:** `200 OK` with the `validationtoken` echoed as `text/plain`.

---

#### Notification Payload

```
POST /api/webhook
Content-Type: application/json
X-SP-Webhook-Signature: <base64-hmac-sha256>  (optional)

{
  "value": [
    {
      "subscriptionId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "clientState": "your-client-state",
      "expirationDateTime": "2026-10-26T20:59:14.000Z",
      "resource": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
      "siteUrl": "/sites/mysite",
      "webId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    }
  ]
}
```

**Success Response:** `200 OK`

**Error Responses:**

| Status | Reason |
|---|---|
| `401 Unauthorized` | clientState mismatch or invalid HMAC signature |
| `500 Internal Server Error` | Unhandled processing error |

---

### GET /api/webhook/health

Returns the health status of the webhook receiver.

```
GET /api/webhook/health
```

**Response:**
```json
{
  "status": "healthy",
  "timestamp": "2026-04-26T20:59:14.000Z"
}
```

---

## 4. Configuration Reference

### `appsettings.json` — Complete schema

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AzureAd": {
    "TenantId":            "<Azure AD Directory (tenant) ID>",
    "ClientId":            "<Azure AD Application (client) ID>",
    "CertificatePath":     "<Relative or absolute path to .pfx file>",
    "CertificatePassword": "<Password protecting the .pfx file>"
  },
  "SharePoint": {
    "SiteUrl":  "<Full SharePoint site collection URL>",
    "ListName": "<Display name of the list>"
  },
  "Webhook": {
    "ClientState":     "<Random secret ≥ 32 chars for notification validation>",
    "NotificationUrl": "<Public HTTPS URL of your /api/webhook endpoint>"
  }
}
```

### Environment variable overrides

Use double underscores (`__`) as section separators:

```bash
AzureAd__TenantId=<value>
AzureAd__ClientId=<value>
AzureAd__CertificatePath=<value>
AzureAd__CertificatePassword=<value>
SharePoint__SiteUrl=<value>
SharePoint__ListName=<value>
Webhook__ClientState=<value>
Webhook__NotificationUrl=<value>
```

---

## 5. Error Codes

| Code | Source | Meaning |
|---|---|---|
| `AADSTS700027` | Azure AD | Certificate signature invalid — re-upload `.cer` |
| `AADSTS7000215` | Azure AD | Wrong credential type — use certificate, not secret |
| `AADSTS65001` | Azure AD | Admin consent required |
| `AADSTS90002` | Azure AD | Tenant not found — verify `TenantId` |
| `HTTP 400` | SharePoint | Bad subscription request — check URL and expiry format |
| `HTTP 401` | SharePoint / API | Invalid token or clientState mismatch |
| `HTTP 403` | SharePoint | Insufficient permissions |
| `HTTP 404` | SharePoint | List or site not found |
