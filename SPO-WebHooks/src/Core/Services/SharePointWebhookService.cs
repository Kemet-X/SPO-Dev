using Microsoft.Extensions.Logging;
using SPO.Webhooks.Core.Models;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace SPO.Webhooks.Core.Services;

/// <summary>
/// Manages SharePoint Online webhook subscriptions using the SharePoint REST API.
/// </summary>
public class SharePointWebhookService
{
    private readonly HttpClient _httpClient;
    private readonly string _sharePointSiteUrl;
    private readonly string _listName;
    private readonly ILogger<SharePointWebhookService> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    /// <summary>
    /// Initializes a new instance of <see cref="SharePointWebhookService"/>.
    /// </summary>
    /// <param name="httpClient">HTTP client used for REST API calls.</param>
    /// <param name="sharePointSiteUrl">Full URL of the SharePoint site collection, e.g. <c>https://contoso.sharepoint.com/sites/mysite</c>.</param>
    /// <param name="listName">Display name of the SharePoint list to manage webhooks for.</param>
    /// <param name="logger">Logger instance.</param>
    public SharePointWebhookService(
        HttpClient httpClient,
        string sharePointSiteUrl,
        string listName,
        ILogger<SharePointWebhookService> logger)
    {
        _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        _sharePointSiteUrl = sharePointSiteUrl?.TrimEnd('/') ?? throw new ArgumentNullException(nameof(sharePointSiteUrl));
        _listName = listName ?? throw new ArgumentNullException(nameof(listName));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Sets the Bearer token on the underlying HTTP client.
    /// </summary>
    /// <param name="accessToken">OAuth 2.0 Bearer token.</param>
    public void SetAccessToken(string accessToken)
    {
        if (string.IsNullOrWhiteSpace(accessToken))
            throw new ArgumentException("Access token must not be null or empty.", nameof(accessToken));

        _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", accessToken);
    }

    /// <summary>
    /// Creates a new webhook subscription on the configured SharePoint list.
    /// </summary>
    /// <param name="notificationUrl">HTTPS URL of the webhook receiver endpoint.</param>
    /// <param name="clientState">Optional secret used to validate incoming notifications via HMAC-SHA256.</param>
    /// <param name="expirationInDays">Number of days until the subscription expires (max 180).</param>
    /// <returns>The created <see cref="WebhookSubscription"/>.</returns>
    /// <exception cref="HttpRequestException">Thrown when the REST API returns a non-success status code.</exception>
    public async Task<WebhookSubscription> CreateSubscriptionAsync(
        string notificationUrl,
        string clientState = "",
        int expirationInDays = 180)
    {
        if (string.IsNullOrWhiteSpace(notificationUrl))
            throw new ArgumentException("Notification URL must not be null or empty.", nameof(notificationUrl));

        if (expirationInDays < 1 || expirationInDays > 180)
            throw new ArgumentOutOfRangeException(nameof(expirationInDays), "Expiration must be between 1 and 180 days.");

        _logger.LogInformation("Creating webhook subscription for list '{ListName}' with notification URL {NotificationUrl}.", _listName, notificationUrl);

        var expiration = DateTime.UtcNow.AddDays(expirationInDays).ToString("o");
        var payload = new
        {
            resource = GetListResourceUrl(),
            notificationUrl,
            expirationDateTime = expiration,
            clientState
        };

        var json = JsonSerializer.Serialize(payload, JsonOptions);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        var requestUrl = GetSubscriptionsUrl();
        var response = await _httpClient.PostAsync(requestUrl, content).ConfigureAwait(false);

        await EnsureSuccessAsync(response, "create subscription").ConfigureAwait(false);

        var responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        var subscription = JsonSerializer.Deserialize<WebhookSubscription>(responseJson, JsonOptions)
            ?? throw new InvalidOperationException("Received an empty subscription response from SharePoint.");

        _logger.LogInformation("Webhook subscription created successfully. ID: {SubscriptionId}", subscription.Id);
        return subscription;
    }

    /// <summary>
    /// Retrieves all webhook subscriptions for the configured SharePoint list.
    /// </summary>
    /// <returns>A list of <see cref="WebhookSubscription"/> objects.</returns>
    public async Task<List<WebhookSubscription>> GetSubscriptionsAsync()
    {
        _logger.LogInformation("Retrieving webhook subscriptions for list '{ListName}'.", _listName);

        var requestUrl = GetSubscriptionsUrl();
        var response = await _httpClient.GetAsync(requestUrl).ConfigureAwait(false);

        await EnsureSuccessAsync(response, "get subscriptions").ConfigureAwait(false);

        var responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

        using var doc = JsonDocument.Parse(responseJson);
        var valuesJson = doc.RootElement.GetProperty("value").GetRawText();

        var subscriptions = JsonSerializer.Deserialize<List<WebhookSubscription>>(valuesJson, JsonOptions)
            ?? new List<WebhookSubscription>();

        _logger.LogInformation("Found {Count} webhook subscription(s) for list '{ListName}'.", subscriptions.Count, _listName);
        return subscriptions;
    }

    /// <summary>
    /// Updates the expiration date of an existing webhook subscription.
    /// </summary>
    /// <param name="subscriptionId">ID of the subscription to update.</param>
    /// <param name="expirationInDays">New number of days until the subscription expires (max 180).</param>
    /// <returns>The updated <see cref="WebhookSubscription"/>.</returns>
    public async Task<WebhookSubscription> RenewSubscriptionAsync(string subscriptionId, int expirationInDays = 180)
    {
        if (string.IsNullOrWhiteSpace(subscriptionId))
            throw new ArgumentException("Subscription ID must not be null or empty.", nameof(subscriptionId));

        if (expirationInDays < 1 || expirationInDays > 180)
            throw new ArgumentOutOfRangeException(nameof(expirationInDays), "Expiration must be between 1 and 180 days.");

        _logger.LogInformation("Renewing subscription {SubscriptionId} by {Days} days.", subscriptionId, expirationInDays);

        var expiration = DateTime.UtcNow.AddDays(expirationInDays).ToString("o");
        var payload = new { expirationDateTime = expiration };

        var json = JsonSerializer.Serialize(payload, JsonOptions);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");

        var requestUrl = $"{GetSubscriptionsUrl()}('{subscriptionId}')";
        var request = new HttpRequestMessage(HttpMethod.Patch, requestUrl) { Content = content };
        var response = await _httpClient.SendAsync(request).ConfigureAwait(false);

        await EnsureSuccessAsync(response, "renew subscription").ConfigureAwait(false);

        var responseJson = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        var subscription = JsonSerializer.Deserialize<WebhookSubscription>(responseJson, JsonOptions)
            ?? throw new InvalidOperationException("Received an empty response when renewing subscription.");

        _logger.LogInformation("Subscription {SubscriptionId} renewed. New expiry: {Expiry}", subscriptionId, subscription.ExpirationDateTime);
        return subscription;
    }

    /// <summary>
    /// Deletes a webhook subscription by ID.
    /// </summary>
    /// <param name="subscriptionId">ID of the subscription to delete.</param>
    public async Task DeleteSubscriptionAsync(string subscriptionId)
    {
        if (string.IsNullOrWhiteSpace(subscriptionId))
            throw new ArgumentException("Subscription ID must not be null or empty.", nameof(subscriptionId));

        _logger.LogInformation("Deleting subscription {SubscriptionId}.", subscriptionId);

        var requestUrl = $"{GetSubscriptionsUrl()}('{subscriptionId}')";
        var response = await _httpClient.DeleteAsync(requestUrl).ConfigureAwait(false);

        await EnsureSuccessAsync(response, "delete subscription").ConfigureAwait(false);

        _logger.LogInformation("Subscription {SubscriptionId} deleted successfully.", subscriptionId);
    }

    // ── Private helpers ─────────────────────────────────────────────────────────

    private string GetListResourceUrl() =>
        $"{_sharePointSiteUrl}/_api/Web/Lists/getByTitle('{_listName}')";

    private string GetSubscriptionsUrl() =>
        $"{GetListResourceUrl()}/Subscriptions";

    private async Task EnsureSuccessAsync(HttpResponseMessage response, string operation)
    {
        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            _logger.LogError("SharePoint REST API error during '{Operation}'. Status: {StatusCode}. Body: {Body}",
                operation, (int)response.StatusCode, body);
            response.EnsureSuccessStatusCode(); // throws HttpRequestException with status code
        }
    }
}
