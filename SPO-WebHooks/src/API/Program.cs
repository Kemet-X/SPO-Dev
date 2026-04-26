using Microsoft.Extensions.Logging;
using SPO.Webhooks.Core.Services;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;

// ── Host configuration ────────────────────────────────────────────────────────
var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new()
    {
        Title = "SPO-WebHooks API",
        Version = "v1",
        Description = "ASP.NET Core API for receiving SharePoint Online webhook notifications using Azure AD certificate authentication."
    });

    // Include XML documentation comments
    var xmlFile = $"{System.Reflection.Assembly.GetExecutingAssembly().GetName().Name}.xml";
    var xmlPath = Path.Combine(AppContext.BaseDirectory, xmlFile);
    if (File.Exists(xmlPath))
        options.IncludeXmlComments(xmlPath);
});

// ── Configuration ─────────────────────────────────────────────────────────────
var configuration = builder.Configuration;

// ── Azure AD authentication service ──────────────────────────────────────────
builder.Services.AddSingleton<AzureAdAuthenticationService>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<AzureAdAuthenticationService>>();

    var tenantId     = configuration["AzureAd:TenantId"]
        ?? throw new InvalidOperationException("AzureAd:TenantId is not configured.");
    var clientId     = configuration["AzureAd:ClientId"]
        ?? throw new InvalidOperationException("AzureAd:ClientId is not configured.");
    var certPath     = configuration["AzureAd:CertificatePath"]
        ?? throw new InvalidOperationException("AzureAd:CertificatePath is not configured.");
    var certPassword = configuration["AzureAd:CertificatePassword"] ?? string.Empty;

    var certificate = new X509Certificate2(certPath, certPassword,
        X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);

    return new AzureAdAuthenticationService(tenantId, clientId, certificate, logger);
});

// ── Webhook notification handler ──────────────────────────────────────────────
builder.Services.AddSingleton<WebhookNotificationHandler>(sp =>
{
    var logger      = sp.GetRequiredService<ILogger<WebhookNotificationHandler>>();
    var clientState = configuration["Webhook:ClientState"] ?? string.Empty;
    return new WebhookNotificationHandler(clientState, logger);
});

// ── HTTP client ───────────────────────────────────────────────────────────────
builder.Services.AddHttpClient();

// ── Build and configure pipeline ──────────────────────────────────────────────
var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(options =>
    {
        options.SwaggerEndpoint("/swagger/v1/swagger.json", "SPO-WebHooks API v1");
        options.RoutePrefix = string.Empty; // Serve Swagger UI at root
    });
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

await app.RunAsync();
