# Troubleshooting Guide

Common issues and their resolutions when running a SharePoint webhook Azure Function.

---

## Table of Contents

- [Authentication Errors](#authentication-errors)
- [SharePoint API Errors](#sharepoint-api-errors)
- [Webhook Registration Errors](#webhook-registration-errors)
- [Webhook Notifications Not Arriving](#webhook-notifications-not-arriving)
- [Azure Function Errors](#azure-function-errors)
- [Network / Connectivity](#network--connectivity)
- [Certificate Issues](#certificate-issues)
- [Diagnostic Script](#diagnostic-script)

---

## Authentication Errors

### `AADSTS700016` – Application not found

**Symptom:** Token request returns `AADSTS700016: Application with identifier '...' was not found in the directory`.

**Causes & Fixes:**
- Wrong `ClientId` – verify in **Azure Portal › App Registrations**.
- Wrong `TenantId` – the app must be registered in the same tenant as SharePoint.
- App was deleted – re-create the registration.

---

### `AADSTS7000215` – Invalid client secret

**Symptom:** `AADSTS7000215: Invalid client secret provided`.

**Fixes:**
1. Verify the secret value was copied completely (no trailing spaces).
2. Check that the secret has not expired (**App Registrations › Certificates & secrets**).
3. Rotate the secret and update `CLIENT_SECRET` / Azure Function App Settings.

---

### `AADSTS65001` – Consent required

**Symptom:** `AADSTS65001: The user or administrator has not consented to use the application`.

**Fix:**
1. Go to **Azure Portal › App Registrations › [your app] › API permissions**.
2. Click **Grant admin consent for [tenant]**.
3. Ensure the consent button shows a green tick next to each permission.

---

### `401 Unauthorized` from SharePoint

**Symptom:** SharePoint REST API returns HTTP 401.

**Fixes (in order):**
1. Confirm the token audience is the SharePoint tenant root (`https://<tenant>.sharepoint.com/`), **not** Graph.
2. Check that the app has `Sites.Manage.All` (or `Sites.Selected`) under **SharePoint API** permissions, not just Graph.
3. If using `Sites.Selected`, ensure the app was explicitly granted access to the site – see [WEBHOOK-REGISTRATION.md](WEBHOOK-REGISTRATION.md#sites-selected-permission).
4. Re-run `Authenticate-SharePointAppIdentity.ps1` and inspect the token audience with a JWT debugger (<https://jwt.ms>).

---

### `403 Forbidden` from SharePoint

**Symptom:** HTTP 403 on certain API endpoints despite having a valid token.

**Fixes:**
- `Sites.Manage.All` is required to create/delete webhook subscriptions (not just `Sites.Read.All`).
- The service principal may need **Site Collection Administrator** rights on the specific site – add via PnP:
  ```powershell
  Add-PnPSiteCollectionAdmin -Owners "i:0i.t|ms.sp.int|<client-id>@<tenant-id>"
  ```

---

### Managed Identity token request fails (HTTP 400 / no IMDS)

**Symptom:** `Invoke-RestMethod` to `http://169.254.169.254/metadata/...` fails or times out.

**Causes & Fixes:**
- Script is not running inside an Azure resource (VM, App Service, Function). The IMDS endpoint is only reachable from within Azure.
- System-assigned identity is not enabled – enable it in **Function App › Identity › System assigned**.
- User-assigned identity – ensure the identity is attached to the Function App and the correct `ClientId` is supplied.

---

## SharePoint API Errors

### `The object specified does not belong to a list` (400)

**Cause:** The list ID or name is incorrect.

**Fix:** Retrieve the correct list ID:
```powershell
$siteUrl  = "https://<tenant>.sharepoint.com/sites/<site>"
$listName = "My List"
$headers  = @{ Authorization = "Bearer $token"; Accept = "application/json;odata=nometadata" }
$r = Invoke-RestMethod -Uri "$siteUrl/_api/web/lists/getbytitle('$listName')?`$select=Id" -Headers $headers
$r.Id
```

---

### `The attempted operation is prohibited because it exceeds the list view threshold` (429 / 500)

**Cause:** Query returns more than 5,000 items.

**Fix:** Add `$top` and `$filter` parameters to reduce the result set.

---

## Webhook Registration Errors

### `notificationUrl is not valid` (400)

**Causes & Fixes:**
- URL must be publicly reachable over HTTPS – SharePoint cannot call `localhost`.
- URL must respond to the validation handshake within **5 seconds** (returns `validationToken` as plain text, HTTP 200).
- Test locally with [ngrok](https://ngrok.com) or [Azure Dev Tunnels](https://learn.microsoft.com/azure/developer/dev-tunnels/).

### `Subscription validation request failed` (400)

SharePoint sends a GET request with `?validationToken=<token>` immediately when you register the webhook.
Your function **must**:
1. Detect the `validationToken` query parameter.
2. Return it as plain text in the response body (not JSON-encoded).
3. Return HTTP 200 within 5 seconds.

Example check (C#):
```csharp
if (req.Query.ContainsKey("validationToken"))
{
    return new OkObjectResult(req.Query["validationToken"].ToString());
}
```

### `Subscription expiration cannot exceed 6 months` (400)

**Fix:** Set `expirationDateTime` to at most 180 days in the future:
```powershell
(Get-Date).AddDays(180).ToUniversalTime().ToString("o")
```

Implement a renewal timer function to automatically extend subscriptions before they expire.

---

## Webhook Notifications Not Arriving

### Checklist

| # | Check | How to verify |
|---|-------|---------------|
| 1 | Subscription exists | `GET /_api/web/lists('<id>')/subscriptions` |
| 2 | Subscription not expired | Check `expirationDateTime` |
| 3 | Notification URL is publicly reachable | Test from external network / Postman |
| 4 | Function returns HTTP 200 within 5 s | Run `Test-SharePointAuth.ps1 -TestSuite Function` |
| 5 | List item was actually changed | Confirm change in SharePoint |
| 6 | Correct list is subscribed | Verify `listId` matches |

### SharePoint retries

SharePoint retries failed notifications up to **5 times** with exponential back-off. If all retries fail, the subscription may be automatically deleted. Monitor your Function App logs in Application Insights.

---

## Azure Function Errors

### Function returns HTTP 500

1. Check logs in **Azure Portal › Function App › Monitor › Logs**.
2. Enable Application Insights for detailed traces.
3. Increase logging in `host.json`:
   ```json
   { "logging": { "logLevel": { "default": "Debug" } } }
   ```

### Cold start timeout

Azure Functions on Consumption plan may cold-start slowly, causing the 5-second validation timeout.

**Fix:** Use **Premium plan** or enable **Always On** (Basic tier or higher App Service plan).

### Function key not in URL

SharePoint notifications do not automatically include the function key. The notification URL registered in SharePoint must include `?code=<function-key>`.

---

## Network / Connectivity

### DNS resolution failure

```powershell
Resolve-DnsName login.microsoftonline.com
Resolve-DnsName <tenant>.sharepoint.com
```

### Outbound firewall blocking Azure AD

Ensure outbound HTTPS (port 443) is allowed to:
- `login.microsoftonline.com`
- `graph.microsoft.com`
- `<tenant>.sharepoint.com`

### Inbound firewall blocking SharePoint

SharePoint webhook callbacks originate from Microsoft data centre IP ranges. Ensure inbound HTTPS is not restricted to known IPs, or add Microsoft's service tags (`AzureCloud`) to your Network Security Group.

---

## Certificate Issues

### `No credentials are available in the security package` or private key error

**Fix:** Load the certificate with `MachineKeySet | PersistKeySet` flags:
```powershell
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
    $path, $password,
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet -bor
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
)
```

### Certificate thumbprint mismatch

Azure AD stores the public certificate. Ensure the thumbprint in your app settings matches:
```powershell
(Get-Item Cert:\CurrentUser\My\<thumbprint>).Thumbprint
```

---

## Diagnostic Script

Run the full test suite for a quick health check:

```powershell
.\scripts\Test-SharePointAuth.ps1 `
    -TenantId          "<tenant-id>" `
    -ClientId          "<client-id>" `
    -ClientSecret      "<secret>" `
    -SharePointSiteUrl "https://<tenant>.sharepoint.com/sites/<site>" `
    -SharePointListName "My List" `
    -FunctionUrl       "https://<func>.azurewebsites.net/api/WebhookTrigger?code=<key>" `
    -TestSuite         "All"
```

The script outputs a **PASS / FAIL / WARN** result for each check and prints actionable messages for failures.
