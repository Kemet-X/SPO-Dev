using System.Text.Json.Serialization;

namespace SPO.Webhooks.Core.Models;

/// <summary>
/// Webhook notification payload received from SharePoint
/// </summary>
public class WebhookNotification
{
    [JsonPropertyName("value")]
    public List<NotificationData> Value { get; set; } = new();
}

public class NotificationData
{
    [JsonPropertyName("subscriptionId")]
    public string SubscriptionId { get; set; } = string.Empty;

    [JsonPropertyName("clientState")]
    public string ClientState { get; set; } = string.Empty;

    [JsonPropertyName("expirationDateTime")]
    public DateTime ExpirationDateTime { get; set; }

    [JsonPropertyName("resource")]
    public string Resource { get; set; } = string.Empty;

    [JsonPropertyName("tenantId")]
    public string TenantId { get; set; } = string.Empty;

    [JsonPropertyName("siteUrl")]
    public string SiteUrl { get; set; } = string.Empty;

    [JsonPropertyName("webId")]
    public string WebId { get; set; } = string.Empty;
}

/// <summary>
/// Subscription validation token request from SharePoint
/// </summary>
public class SubscriptionValidationRequest
{
    [JsonPropertyName("value")]
    public string ValidationToken { get; set; } = string.Empty;
}