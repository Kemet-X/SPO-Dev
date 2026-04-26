# Troubleshooting Guide — SPO-WebHooks

This guide covers the most common issues encountered when setting up and running the SharePoint Online Webhooks project with Azure AD certificate authentication.

---

## Table of Contents

1. [Certificate Issues](#1-certificate-issues)
2. [Azure AD Authentication Failures](#2-azure-ad-authentication-failures)
3. [SharePoint Webhook Subscription Issues](#3-sharepoint-webhook-subscription-issues)
4. [Webhook Notification Reception Issues](#4-webhook-notification-reception-issues)
5. [Deployment & Runtime Issues](#5-deployment--runtime-issues)
6. [Debugging Tips](#6-debugging-tips)
7. [Common Error Messages](#7-common-error-messages)
8. [Quick Checklist for New Setup](#8-quick-checklist-for-new-setup)
9. [Getting Help](#9-getting-help)

---

## 1. Certificate Issues

### Problem: "No credentials are available in the security package"

**Root cause:** The certificate file cannot be loaded because the path is wrong, the password is incorrect, or the key storage flags are incompatible with the runtime environment.

**Solution:**
```csharp
// Use MachineKeySet when running as a Windows Service or in IIS
var certificate = new X509Certificate2(
    certPath,
    certPassword,
    X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet
);
```

For Linux/macOS containers use `EphemeralKeySet` instead:
```csharp
X509KeyStorageFlags.EphemeralKeySet
```

---

### Problem: Certificate thumbprint mismatch

**Root cause:** The thumbprint stored in Azure AD differs from the certificate loaded at runtime.

**Solution — Get thumbprint via PowerShell:**
```powershell
$cert = Get-PfxCertificate -FilePath ".\app-cert.pfx"
Write-Host "Thumbprint: $($cert.Thumbprint)"
```

**Solution — Verify in C#:**
```csharp
var cert = new X509Certificate2("app-cert.pfx", "password");
Console.WriteLine($"Thumbprint: {cert.Thumbprint}");
```

Compare the output with the thumbprint shown in **Azure AD → App registrations → Certificates & secrets**.

---

### Problem: Certificate is expired

**Check expiry:**
```powershell
$cert = Get-PfxCertificate -FilePath ".\app-cert.pfx"
Write-Host "Expires: $($cert.NotAfter)"
```

**Renew a self-signed certificate:**
```powershell
$cert = New-SelfSignedCertificate `
    -DnsName "SPO-WebHooks" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -NotAfter (Get-Date).AddYears(2)

$pwd = ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath ".\app-cert.pfx" -Password $pwd
```

---

## 2. Azure AD Authentication Failures

### Problem: `AADSTS700027` — Client assertion contains an invalid signature

**Root cause:** The certificate uploaded to Azure AD does not match the certificate being used at runtime.

**Solution:**
1. Re-export the `.cer` (public key only) from your `.pfx`:
   ```powershell
   $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(".\app-cert.pfx", "pwd")
   [IO.File]::WriteAllBytes(".\app-cert.cer", $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
   ```
2. Upload the `.cer` file to **Azure AD → App registrations → Certificates & secrets → Upload certificate**.

---

### Problem: `AADSTS7000215` — Invalid client secret

**Root cause:** You are using a client secret instead of a certificate, or the certificate is being passed incorrectly.

**Solution:** Remove any `WithClientSecret()` call and ensure you are using `WithCertificate()`:
```csharp
var app = ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithAuthority($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token")
    .WithCertificate(certificate)   // ← certificate, NOT secret
    .Build();
```

---

### Problem: `AADSTS65001` — User has not consented to the application

**Root cause:** Admin consent has not been granted for the required SharePoint API permissions.

**Solution:**
1. In **Azure AD → App registrations → API permissions**, add `Sites.Read.All` (Application permission).
2. Click **Grant admin consent for \<tenant\>**.
3. Confirm all permissions show a ✅ green status.

---

### Problem: Token scope is wrong — SharePoint returns 401

**Root cause:** The token was acquired for Microsoft Graph but is being used for SharePoint REST API.

**Solution:** Use a SharePoint-specific scope:
```csharp
// Wrong — Graph scope
var scopes = new[] { "https://graph.microsoft.com/.default" };

// Correct — SharePoint scope
var scopes = new[] { "https://<your-tenant>.sharepoint.com/.default" };
```

---

## 3. SharePoint Webhook Subscription Issues

### Problem: 400 Bad Request when creating a subscription

**Common causes and fixes:**

| Cause | Fix |
|---|---|
| `notificationUrl` is HTTP, not HTTPS | Use an HTTPS URL |
| `notificationUrl` is not publicly reachable | Use ngrok for local dev: `ngrok http 5001` |
| `expirationDateTime` format is wrong | Use ISO 8601: `DateTime.UtcNow.AddDays(180).ToString("o")` |
| List name has a typo | Double-check in SharePoint admin |

---

### Problem: Subscription validation handshake fails

**Root cause:** SharePoint sends a GET with `?validationtoken=<token>` and expects the token echoed back as plain text within 5 seconds.

**Solution — correct ASP.NET Core handler:**
```csharp
[HttpPost]
public async Task<IActionResult> ReceiveNotification([FromQuery] string? validationToken = null)
{
    if (!string.IsNullOrEmpty(validationToken))
        return Content(validationToken, "text/plain");  // ← must be plain text, not JSON

    // ... process notification
}
```

**For local development** use [ngrok](https://ngrok.com/) to expose your localhost:
```bash
ngrok http https://localhost:5001
# Use the https://xxxx.ngrok.io URL as your notificationUrl
```

---

### Problem: Subscription expires after 180 days

SharePoint webhook subscriptions expire after a maximum of 180 days. Implement automatic renewal:

```csharp
// Renew before it expires
if (subscription.ExpirationDateTime < DateTime.UtcNow.AddDays(7))
{
    await webhookService.RenewSubscriptionAsync(subscription.Id, expirationInDays: 180);
}
```

Schedule this check to run weekly via Azure Functions, a cron job, or a background service.

---

### Problem: 403 Forbidden when calling the Subscriptions API

**Root cause:** The Azure AD application lacks the required SharePoint permissions, or admin consent was not granted.

**Required permissions (application-level):**
- `Sites.Read.All` — Read items in all site collections
- `Sites.Manage.All` — Create/manage webhook subscriptions

---

## 4. Webhook Notification Reception Issues

### Problem: Notifications are not received

**Checklist:**
- [ ] The `notificationUrl` is publicly accessible over HTTPS
- [ ] Firewall / NSG / WAF rules allow inbound POST requests from SharePoint IPs
- [ ] The subscription is not expired (`ExpirationDateTime` > now)
- [ ] The list actually has changes (create/update/delete an item to trigger)
- [ ] The app is running and listening on the correct port

**Test reachability:**
```bash
curl -X POST https://your-api-domain/api/webhook \
  -H "Content-Type: application/json" \
  -d '{"value":[]}'
```

---

### Problem: `clientState` mismatch — notifications rejected

**Root cause:** The `clientState` in the incoming notification does not match the value used when creating the subscription.

**Solution:** Ensure the same `clientState` value is used in both places:
```json
// appsettings.json
{
  "Webhook": {
    "ClientState": "YourSecretHere"   // ← must match subscription creation
  }
}
```

**Generate a strong clientState:**
```csharp
var clientState = Convert.ToBase64String(
    System.Security.Cryptography.RandomNumberGenerator.GetBytes(32)
);
```

---

## 5. Deployment & Runtime Issues

### Problem: Application crashes on startup in Docker / Azure App Service

**Root cause:** Certificate loading fails because `MachineKeySet` is not available on Linux.

**Fix:** Use `EphemeralKeySet` on Linux:
```csharp
var flags = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
    ? X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet
    : X509KeyStorageFlags.EphemeralKeySet;

var certificate = new X509Certificate2(certPath, certPassword, flags);
```

---

### Problem: Configuration values are null at runtime

**Root cause:** `appsettings.json` is not copied to the output directory, or the key names don't match.

**Fix — ensure file is copied:**
```xml
<!-- In .csproj -->
<None Update="appsettings.json">
  <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
</None>
```

**Use environment variables** to override config in production:
```bash
export AzureAd__TenantId="your-tenant-id"
export AzureAd__ClientId="your-client-id"
```
Note: ASP.NET Core uses `__` (double underscore) as the hierarchy separator for environment variables.

---

### Problem: `HttpClient` SSL errors in development

**Fix:**
```bash
dotnet dev-certs https --trust
```

---

## 6. Debugging Tips

### Enable verbose MSAL logging

```csharp
var app = ConfidentialClientApplicationBuilder
    .Create(clientId)
    .WithAuthority(authorityUrl)
    .WithCertificate(certificate)
    .WithLogging((level, message, _) =>
    {
        Console.WriteLine($"[MSAL {level}] {message}");
    }, LogLevel.Verbose, enablePiiLogging: false)
    .Build();
```

### Test your webhook endpoint locally with ngrok

```bash
# Install ngrok from https://ngrok.com/
ngrok http https://localhost:5001

# Your public URL will be something like:
# https://abc123.ngrok.io -> https://localhost:5001
```

### Check SharePoint webhook subscriptions via REST

```bash
curl -X GET \
  "https://<tenant>.sharepoint.com/sites/<site>/_api/Web/Lists/getByTitle('<list>')/Subscriptions" \
  -H "Authorization: Bearer <token>" \
  -H "Accept: application/json;odata=nometadata"
```

### Decode a JWT access token

Paste your token at [jwt.ms](https://jwt.ms) to inspect claims, expiry, and scope.

---

## 7. Common Error Messages

| Error Code | Message | Resolution |
|---|---|---|
| `AADSTS700027` | Invalid certificate signature | Re-upload `.cer` to Azure AD |
| `AADSTS7000215` | Invalid client secret | Switch from secret to certificate |
| `AADSTS65001` | Consent required | Grant admin consent in Azure AD |
| `AADSTS90002` | Tenant not found | Verify `TenantId` in config |
| `AADSTS70011` | Invalid scope | Use `<sharepoint-url>/.default` scope |
| `HTTP 400` | Bad request creating subscription | Check `notificationUrl` is HTTPS and reachable |
| `HTTP 401` | Unauthorized from SharePoint | Token scope or permissions issue |
| `HTTP 403` | Forbidden | Grant `Sites.Manage.All` and admin consent |
| `HTTP 404` | List not found | Check list name and site URL |

---

## 8. Quick Checklist for New Setup

- [ ] Azure AD app registration created
- [ ] Certificate generated (`.pfx` with private key, `.cer` for upload)
- [ ] `.cer` uploaded to Azure AD app → **Certificates & secrets**
- [ ] `Sites.Read.All` Application permission added
- [ ] `Sites.Manage.All` Application permission added
- [ ] Admin consent granted for all permissions
- [ ] `TenantId`, `ClientId`, `CertificatePath`, `CertificatePassword` set in `appsettings.json`
- [ ] SharePoint site URL and list name set correctly
- [ ] Webhook receiver endpoint is publicly accessible via HTTPS
- [ ] `notificationUrl` is set to your HTTPS endpoint
- [ ] `clientState` is a strong random secret (≥ 32 characters)
- [ ] Validation handshake handler echoes `validationToken` as plain text
- [ ] Subscription expiry renewal is scheduled
- [ ] Certificates are **not** committed to source control (check `.gitignore`)
- [ ] Sensitive values are stored in Azure Key Vault or environment variables in production

---

## 9. Getting Help

- **Microsoft Docs — SharePoint Webhooks:** https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/overview-sharepoint-webhooks
- **Microsoft Docs — MSAL.NET:** https://learn.microsoft.com/en-us/entra/msal/dotnet/
- **Project Issues:** Open an issue in the [GitHub repository](https://github.com/Kemet-X/SPO-Dev/issues)
- **LinkedIn:** [H. Hegab](https://www.linkedin.com/in/hhegab/)
