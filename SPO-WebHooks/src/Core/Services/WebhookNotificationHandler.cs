using Microsoft.Extensions.Logging;
using SPO.Webhooks.Core.Models;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace SPO.Webhooks.Core.Services;

/// <summary>
/// Processes and validates incoming SharePoint webhook notifications.
/// Verifies HMAC-SHA256 signatures when a <c>clientState</c> secret is configured.
/// </summary>
public class WebhookNotificationHandler
{
    private readonly string _clientState;
    private readonly ILogger<WebhookNotificationHandler> _logger;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    /// <summary>
    /// Initializes a new instance of <see cref="WebhookNotificationHandler"/>.
    /// </summary>
    /// <param name="clientState">
    /// The secret string set when the webhook subscription was created.
    /// Used to verify incoming notification signatures. May be empty to skip signature validation.
    /// </param>
    /// <param name="logger">Logger instance.</param>
    public WebhookNotificationHandler(string clientState, ILogger<WebhookNotificationHandler> logger)
    {
        _clientState = clientState ?? string.Empty;
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Validates an incoming webhook notification payload.
    /// Checks that the <c>clientState</c> in each notification entry matches the configured secret
    /// and optionally verifies the HMAC-SHA256 signature header.
    /// </summary>
    /// <param name="payload">Raw JSON body of the incoming POST request.</param>
    /// <param name="signatureHeader">Value of the <c>X-SP-Webhook-Signature</c> header, if present.</param>
    /// <returns><c>true</c> if the notification is valid; otherwise <c>false</c>.</returns>
    public bool ValidateNotification(string payload, string? signatureHeader = null)
    {
        if (string.IsNullOrWhiteSpace(payload))
        {
            _logger.LogWarning("Received empty notification payload.");
            return false;
        }

        // Verify HMAC-SHA256 signature when both a clientState and a signature header are present
        if (!string.IsNullOrEmpty(_clientState) && !string.IsNullOrEmpty(signatureHeader))
        {
            if (!VerifySignature(payload, signatureHeader))
            {
                _logger.LogWarning("Webhook signature validation failed.");
                return false;
            }
        }

        // Verify that each notification entry contains the expected clientState
        try
        {
            var notification = JsonSerializer.Deserialize<WebhookNotification>(payload, JsonOptions);
            if (notification?.Value is null || notification.Value.Count == 0)
            {
                _logger.LogWarning("Notification payload contains no entries.");
                return false;
            }

            foreach (var entry in notification.Value)
            {
                if (!string.IsNullOrEmpty(_clientState) && entry.ClientState != _clientState)
                {
                    _logger.LogWarning(
                        "clientState mismatch for subscription {SubscriptionId}. Expected secret does not match.",
                        entry.SubscriptionId);
                    return false;
                }
            }
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Failed to deserialize notification payload.");
            return false;
        }

        return true;
    }

    /// <summary>
    /// Parses a validated webhook notification payload into a <see cref="WebhookNotification"/> object.
    /// </summary>
    /// <param name="payload">Raw JSON body of the incoming POST request.</param>
    /// <returns>The deserialized <see cref="WebhookNotification"/>.</returns>
    /// <exception cref="JsonException">Thrown when the payload cannot be parsed.</exception>
    public WebhookNotification ParseNotification(string payload)
    {
        if (string.IsNullOrWhiteSpace(payload))
            throw new ArgumentException("Payload must not be null or empty.", nameof(payload));

        var notification = JsonSerializer.Deserialize<WebhookNotification>(payload, JsonOptions)
            ?? throw new InvalidOperationException("Deserialized notification was null.");

        _logger.LogInformation(
            "Parsed webhook notification with {Count} entr(ies). First subscription: {SubscriptionId}",
            notification.Value.Count,
            notification.Value.FirstOrDefault()?.SubscriptionId ?? "N/A");

        return notification;
    }

    /// <summary>
    /// Processes a webhook notification by invoking the provided handler for each notification entry.
    /// </summary>
    /// <param name="notification">The parsed webhook notification.</param>
    /// <param name="handler">Async delegate that processes a single <see cref="NotificationData"/> entry.</param>
    public async Task ProcessNotificationsAsync(
        WebhookNotification notification,
        Func<NotificationData, Task> handler)
    {
        ArgumentNullException.ThrowIfNull(notification);
        ArgumentNullException.ThrowIfNull(handler);

        foreach (var entry in notification.Value)
        {
            _logger.LogInformation(
                "Processing notification for subscription {SubscriptionId}, list {Resource}.",
                entry.SubscriptionId, entry.Resource);

            try
            {
                await handler(entry).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex,
                    "Error processing notification for subscription {SubscriptionId}.",
                    entry.SubscriptionId);
                // Continue processing remaining entries rather than aborting
            }
        }
    }

    // ── Private helpers ──────────────────────────────────────────────────────────

    private bool VerifySignature(string payload, string signatureHeader)
    {
        try
        {
            var keyBytes = Encoding.UTF8.GetBytes(_clientState);
            var payloadBytes = Encoding.UTF8.GetBytes(payload);

            using var hmac = new HMACSHA256(keyBytes);
            var hash = hmac.ComputeHash(payloadBytes);
            var expectedSignature = Convert.ToBase64String(hash);

            var isValid = string.Equals(expectedSignature, signatureHeader, StringComparison.Ordinal);
            if (!isValid)
            {
                // Sanitize the user-supplied header value before logging to prevent log forging
                var sanitized = signatureHeader?.Replace("\r", "").Replace("\n", "") ?? string.Empty;
                _logger.LogDebug(
                    "Signature mismatch. Expected: {Expected}, Received: {Received}",
                    expectedSignature, sanitized);
            }

            return isValid;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Exception during signature verification.");
            return false;
        }
    }
}
