using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System.Security.Cryptography.X509Certificates;

namespace SPO.Webhooks.Core.Services;

/// <summary>
/// Provides Azure AD certificate-based authentication for SharePoint Online access.
/// Uses the OAuth 2.0 client credentials flow with an X.509 certificate.
/// </summary>
public class AzureAdAuthenticationService
{
    private readonly string _tenantId;
    private readonly string _clientId;
    private readonly X509Certificate2 _certificate;
    private readonly ILogger<AzureAdAuthenticationService> _logger;
    private readonly Lazy<IConfidentialClientApplication> _app;

    /// <summary>
    /// Initializes a new instance of <see cref="AzureAdAuthenticationService"/>.
    /// </summary>
    /// <param name="tenantId">Azure AD tenant ID (directory ID).</param>
    /// <param name="clientId">Azure AD application (client) ID.</param>
    /// <param name="certificate">X.509 certificate used for client assertion.</param>
    /// <param name="logger">Logger instance.</param>
    /// <exception cref="ArgumentException">Thrown when <paramref name="tenantId"/> or <paramref name="clientId"/> is null, empty, or whitespace.</exception>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="certificate"/> is null.</exception>
    public AzureAdAuthenticationService(
        string tenantId,
        string clientId,
        X509Certificate2 certificate,
        ILogger<AzureAdAuthenticationService> logger)
    {
        if (string.IsNullOrWhiteSpace(tenantId))
            throw new ArgumentException("Tenant ID must not be null or empty.", nameof(tenantId));

        if (string.IsNullOrWhiteSpace(clientId))
            throw new ArgumentException("Client ID must not be null or empty.", nameof(clientId));

        _certificate = certificate ?? throw new ArgumentNullException(nameof(certificate));
        _tenantId = tenantId;
        _clientId = clientId;
        _logger = logger;

        // Defer MSAL application creation until first use so the constructor
        // succeeds even when the certificate has not yet been fully loaded
        // (e.g. in unit-test scenarios that supply an empty placeholder cert).
        _app = new Lazy<IConfidentialClientApplication>(() =>
            ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithAuthority(GetAuthorityUrl())
                .WithCertificate(certificate)
                .Build());
    }

    /// <summary>
    /// Acquires an OAuth 2.0 access token for the SharePoint Online resource.
    /// </summary>
    /// <returns>A Bearer access token string.</returns>
    /// <exception cref="MsalServiceException">Thrown when Azure AD returns an error response.</exception>
    /// <exception cref="MsalClientException">Thrown when the MSAL client encounters a local error.</exception>
    public async Task<string> GetAccessTokenAsync()
    {
        _logger.LogInformation("Acquiring access token for tenant {TenantId}, client {ClientId}.", _tenantId, _clientId);

        try
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var result = await _app.Value
                .AcquireTokenForClient(scopes)
                .ExecuteAsync()
                .ConfigureAwait(false);

            _logger.LogInformation("Access token acquired successfully. Expires: {ExpiresOn}", result.ExpiresOn);
            return result.AccessToken;
        }
        catch (MsalServiceException ex)
        {
            _logger.LogError(ex, "Azure AD returned an error while acquiring token. Error: {ErrorCode}", ex.ErrorCode);
            throw;
        }
        catch (MsalClientException ex)
        {
            _logger.LogError(ex, "MSAL client error while acquiring token. Error: {ErrorCode}", ex.ErrorCode);
            throw;
        }
    }

    /// <summary>
    /// Acquires an OAuth 2.0 access token scoped to a specific SharePoint tenant.
    /// </summary>
    /// <param name="sharePointTenantUrl">Root URL of the SharePoint tenant, e.g. <c>https://contoso.sharepoint.com</c>.</param>
    /// <returns>A Bearer access token string.</returns>
    public async Task<string> GetSharePointAccessTokenAsync(string sharePointTenantUrl)
    {
        if (string.IsNullOrWhiteSpace(sharePointTenantUrl))
            throw new ArgumentException("SharePoint tenant URL must not be null or empty.", nameof(sharePointTenantUrl));

        _logger.LogInformation("Acquiring SharePoint access token for {SharePointUrl}.", sharePointTenantUrl);

        var trimmed = sharePointTenantUrl.TrimEnd('/');
        var scopes = new[] { $"{trimmed}/.default" };

        var result = await _app.Value
            .AcquireTokenForClient(scopes)
            .ExecuteAsync()
            .ConfigureAwait(false);

        _logger.LogInformation("SharePoint access token acquired. Expires: {ExpiresOn}", result.ExpiresOn);
        return result.AccessToken;
    }

    /// <summary>
    /// Returns the OAuth 2.0 token endpoint URL for the configured tenant.
    /// </summary>
    public string GetAuthorityUrl() =>
        $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
}
