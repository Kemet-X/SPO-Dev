// ====== CONFIG ======
const clientId = "YOUR_CLIENT_ID";
const tenantId = "YOUR_TENANT_ID";

// SharePoint resource (IMPORTANT): token must be for SharePoint, not Graph
const spHostname = "TENANT_NAME.sharepoint.com";         // e.g. contoso.sharepoint.com
const sitePath   = "/sites/SitesSelectedTest";           // e.g. /sites/MySite
const siteUrl    = `https://${spHostname}${sitePath}`;

// Scope for SharePoint delegated token (v2): use resource + /.default (delegated perms must exist on app)
const spScopes = ["openid", "profile", "offline_access", `https://${spHostname}/.default`];

// ====== MSAL INIT ======
const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: "http://localhost:3000/"
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

const output = document.getElementById("output");
function log(msg) { output.textContent += msg + "\n"; }

// Handle redirect response (PKCE exchange happens inside MSAL)
msalInstance.handleRedirectPromise()
  .then((result) => {
    if (result?.account) {
      log("Redirect login completed.");
      log(`Signed in as: ${result.account.username}`);
    }
  })
  .catch((e) => log("handleRedirectPromise error: " + e));

// ====== LOGIN ======
document.getElementById("login").onclick = () => {
  log("Starting loginRedirect (Authorization Code + PKCE)...");
  msalInstance.loginRedirect({ scopes: spScopes });
};

// ====== TOKEN + SHAREPOINT REST CALL ======
document.getElementById("callApi").onclick = async () => {
  try {
    const account = msalInstance.getAllAccounts()[0];
    if (!account) {
      log("No signed-in user. Click Login first.");
      return;
    }

    // Acquire SharePoint access token silently (PKCE already used in initial code exchange)
    const tokenResult = await msalInstance.acquireTokenSilent({
      account,
      scopes: spScopes
    });

    log(" SharePoint access token acquired.");
    log("Access token (first 120 chars): " + tokenResult.accessToken.substring(0, 120) + "...");

    // ---- Call SharePoint REST: GET _api/web/lists ----
    const listsEndpoint = `${siteUrl}/_api/web/lists?$select=Title,Id,Hidden`;

    const resp = await fetch(listsEndpoint, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${tokenResult.accessToken}`,
        "Accept": "application/json;odata=nometadata"
      }
    });

    if (!resp.ok) {
      const text = await resp.text();
      log(`SharePoint REST failed: ${resp.status} ${resp.statusText}`);
      log(text);
      return;
    }

    const data = await resp.json();
    log("SharePoint REST lists response received.");
    log(JSON.stringify(data, null, 2));

  } catch (e) {
    log(" Error: " + (e?.message || e));
    log(JSON.stringify(e, null, 2));
  }
};
