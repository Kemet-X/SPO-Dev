using System.Text.Json.Serialization;

namespace SPO.Webhooks.Core.Models;

/// <summary>
/// Represents a SharePoint webhook subscription
/// </summary>
public class WebhookSubscription
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;

    [JsonPropertyName("resource")]
    public string Resource { get; set; } = string.Empty;

    [JsonPropertyName("notificationUrl")]
    public string NotificationUrl { get; set; } = string.Empty;

    [JsonPropertyName("clientState")]
    public string ClientState { get; set; } = string.Empty;

    [JsonPropertyName("expirationDateTime")]
    public DateTime ExpirationDateTime { get; set; }

    [JsonPropertyName("creationTime")]
    public DateTime CreationTime { get; set; }

    public override string ToString()
    {
        return $"Id: {Id}, Resource: {Resource}, NotificationUrl: {NotificationUrl}, Expires: {ExpirationDateTime:O}";
    }
}