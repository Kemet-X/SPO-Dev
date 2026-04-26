using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using SPO.Webhooks.Core.Services;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

// ── Host setup ────────────────────────────────────────────────────────────────
var host = Host.CreateDefaultBuilder(args)
    .ConfigureAppConfiguration((_, config) =>
    {
        config.SetBasePath(Directory.GetCurrentDirectory())
              .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
              .AddEnvironmentVariables();
    })
    .ConfigureLogging(logging =>
    {
        logging.ClearProviders();
        logging.AddConsole();
    })
    .Build();

var logger      = host.Services.GetRequiredService<ILogger<Program>>();
var config      = host.Services.GetRequiredService<IConfiguration>();

// ── Read configuration ────────────────────────────────────────────────────────
var tenantId         = config["AzureAd:TenantId"]         ?? throw new InvalidOperationException("AzureAd:TenantId is not configured.");
var clientId         = config["AzureAd:ClientId"]         ?? throw new InvalidOperationException("AzureAd:ClientId is not configured.");
var certPath         = config["AzureAd:CertificatePath"]  ?? throw new InvalidOperationException("AzureAd:CertificatePath is not configured.");
var certPassword     = config["AzureAd:CertificatePassword"] ?? string.Empty;
var sharePointSiteUrl = config["SharePoint:SiteUrl"]      ?? throw new InvalidOperationException("SharePoint:SiteUrl is not configured.");
var listName         = config["SharePoint:ListName"]       ?? throw new InvalidOperationException("SharePoint:ListName is not configured.");
var notificationUrl  = config["Webhook:NotificationUrl"]  ?? throw new InvalidOperationException("Webhook:NotificationUrl is not configured.");
var clientState      = config["Webhook:ClientState"]      ?? GenerateClientState();

logger.LogInformation("=== SPO-WebHooks Console Demo ===");
logger.LogInformation("Tenant: {TenantId} | Client: {ClientId}", tenantId, clientId);

// ── Build services ────────────────────────────────────────────────────────────
var certificate = new X509Certificate2(certPath, certPassword,
    X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);

var authLogger    = host.Services.GetRequiredService<ILogger<AzureAdAuthenticationService>>();
var serviceLogger = host.Services.GetRequiredService<ILogger<SharePointWebhookService>>();

var authService    = new AzureAdAuthenticationService(tenantId, clientId, certificate, authLogger);
var httpClient     = new HttpClient();
var webhookService = new SharePointWebhookService(httpClient, sharePointSiteUrl, listName, serviceLogger);

// ── Authenticate ──────────────────────────────────────────────────────────────
logger.LogInformation("Step 1: Acquiring access token via certificate authentication...");
var accessToken = await authService.GetSharePointAccessTokenAsync(sharePointSiteUrl);
webhookService.SetAccessToken(accessToken);
logger.LogInformation("Access token acquired successfully.");

// ── List existing subscriptions ───────────────────────────────────────────────
logger.LogInformation("Step 2: Listing existing webhook subscriptions for '{ListName}'...", listName);
var subscriptions = await webhookService.GetSubscriptionsAsync();

if (subscriptions.Count == 0)
{
    logger.LogInformation("No existing subscriptions found.");
}
else
{
    foreach (var sub in subscriptions)
        logger.LogInformation("  • {Subscription}", sub);
}

// ── Create a new subscription ─────────────────────────────────────────────────
logger.LogInformation("Step 3: Creating a new webhook subscription...");
logger.LogInformation("  Notification URL : {Url}", notificationUrl);
logger.LogInformation("  ClientState      : (secret — not logged)");
logger.LogInformation("  Expiry           : 180 days");

var newSubscription = await webhookService.CreateSubscriptionAsync(
    notificationUrl: notificationUrl,
    clientState: clientState,
    expirationInDays: 180);

logger.LogInformation("Subscription created: {Subscription}", newSubscription);

// ── Renew the subscription ────────────────────────────────────────────────────
logger.LogInformation("Step 4: Renewing the new subscription by another 180 days...");
var renewed = await webhookService.RenewSubscriptionAsync(newSubscription.Id, expirationInDays: 180);
logger.LogInformation("Renewed subscription: {Subscription}", renewed);

// ── Delete the subscription ───────────────────────────────────────────────────
logger.LogInformation("Step 5: Cleaning up — deleting the subscription...");
await webhookService.DeleteSubscriptionAsync(newSubscription.Id);
logger.LogInformation("Subscription {Id} deleted.", newSubscription.Id);

logger.LogInformation("=== Demo complete ===");

// ── Helper ────────────────────────────────────────────────────────────────────
static string GenerateClientState()
{
    var bytes = RandomNumberGenerator.GetBytes(32);
    return Convert.ToBase64String(bytes);
}
