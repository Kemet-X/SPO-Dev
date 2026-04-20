# SPO-DEV - SharePoint Online Developer Examples

This folder contains C# code examples for creating SharePoint Online lists using **Microsoft Graph API** and **CSOM** with different authentication approaches.

## Files

| File | Description |
|---|---|
| `CreateListWithGraph.cs` | Creates a SharePoint list using **Microsoft Graph** with **app-only** (client credentials) auth |
| `CreateListWithCSOM.cs` | Creates a SharePoint list using **CSOM** with **delegated** (interactive) auth — preserves the actual user as "Created By" |
| `GetListCreatedBy.cs` | Retrieves the "Created By" metadata of a SharePoint list using CSOM |
| `DeviceCodeAuth.cs` | Alternative authentication using device code flow (for environments without a browser) |
| `CreateSharePointList.csproj` | Minimal project file with required NuGet packages |

## Authentication Approaches

### 1. App-Only (Client Credentials) — `CreateListWithGraph.cs`
- Uses `ClientSecretCredential` with client ID & secret
- Requires **Application** permission: `Sites.Selected`
- A Global/SharePoint Admin must grant the app `write` access to the specific site
- **"Created By" will show the app name**, not a user

### 2. Delegated (Interactive) — `CreateListWithCSOM.cs`
- Uses MSAL interactive or device code flow
- Requires **Delegated** permission: `AllSites.Manage`
- **"Created By" will show the actual signed-in user** ✅

## Azure AD App Registration

| Setting | App-Only | Delegated |
|---|---|---|
| Platform | N/A | Mobile and desktop |
| Redirect URI | N/A | `http://localhost` |
| API Permissions | Graph > Application > `Sites.Selected` | SharePoint > Delegated > `AllSites.Manage` |
| Client Secret | Required | Not needed |
| Allow public client flows | N/A | Yes |

## NuGet Packages

```
dotnet add package Microsoft.Graph
dotnet add package Azure.Identity
dotnet add package Microsoft.Identity.Client
dotnet add package Microsoft.SharePointOnline.CSOM
```

## Getting Your Site ID

```csharp
// For https://contoso.sharepoint.com/sites/MyProject
var site = await graphClient.Sites["contoso.sharepoint.com:/sites/MyProject"].GetAsync();
Console.WriteLine($"Site ID: {site?.Id}");
```

## Granting Sites.Selected Permission

```http
POST https://graph.microsoft.com/v1.0/sites/{siteId}/permissions

{
  "roles": ["write"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "YOUR_CLIENT_ID",
        "displayName": "My App"
      }
    }
  ]
}
```