using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using SPO.Webhooks.Core.Models;
using SPO.Webhooks.Core.Services;

namespace SPO.Webhooks.API.Controllers;

/// <summary>
/// Handles incoming SharePoint webhook validation challenges and notification events.
/// </summary>
[ApiController]
[Route("api/[controller]")]
[Produces("application/json")]
public class WebhookController : ControllerBase
{
    private readonly WebhookNotificationHandler _notificationHandler;
    private readonly ILogger<WebhookController> _logger;

    /// <summary>
    /// Initializes a new instance of <see cref="WebhookController"/>.
    /// </summary>
    public WebhookController(
        WebhookNotificationHandler notificationHandler,
        ILogger<WebhookController> logger)
    {
        _notificationHandler = notificationHandler ?? throw new ArgumentNullException(nameof(notificationHandler));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    /// <summary>
    /// Webhook receiver endpoint.
    /// <para>
    /// SharePoint first sends a GET request with a <c>validationtoken</c> query parameter to verify
    /// the endpoint. On successful subscription, SharePoint sends POST requests with notification payloads.
    /// </para>
    /// </summary>
    /// <param name="validationToken">
    /// Validation token sent by SharePoint during subscription creation handshake.
    /// When present the endpoint must echo back the token as plain text with HTTP 200.
    /// </param>
    /// <returns>
    /// <c>200 OK</c> with the raw validation token on a validation request,
    /// <c>200 OK</c> on a successfully processed notification,
    /// or <c>401 Unauthorized</c> when the notification signature is invalid.
    /// </returns>
    [HttpPost]
    public async Task<IActionResult> ReceiveNotification([FromQuery] string? validationToken = null)
    {
        // ── Validation handshake (subscription creation) ─────────────────────────
        if (!string.IsNullOrEmpty(validationToken))
        {
            _logger.LogInformation("Received webhook validation request.");
            return Content(validationToken, "text/plain");
        }

        // ── Notification payload ──────────────────────────────────────────────────
        string payload;
        using (var reader = new StreamReader(Request.Body))
        {
            payload = await reader.ReadToEndAsync().ConfigureAwait(false);
        }

        _logger.LogInformation("Received webhook notification. Payload length: {Length} bytes.", payload.Length);

        var signatureHeader = Request.Headers["X-SP-Webhook-Signature"].FirstOrDefault();

        if (!_notificationHandler.ValidateNotification(payload, signatureHeader))
        {
            _logger.LogWarning("Webhook notification validation failed. Returning 401.");
            return Unauthorized(new { error = "Invalid webhook signature or clientState mismatch." });
        }

        var notification = _notificationHandler.ParseNotification(payload);

        // Process each notification asynchronously.
        // SharePoint expects a 200 response within 5 seconds; for long-running work
        // enqueue the entries and return immediately.
        await _notificationHandler.ProcessNotificationsAsync(notification, async entry =>
        {
            _logger.LogInformation(
                "Processing change on site {SiteUrl}, list {Resource}, subscription {SubscriptionId}.",
                entry.SiteUrl, entry.Resource, entry.SubscriptionId);

            // TODO: Replace this placeholder with your business logic,
            //       e.g. queuing the entry to Azure Service Bus or Azure Queue Storage.
            await Task.CompletedTask.ConfigureAwait(false);
        }).ConfigureAwait(false);

        return Ok();
    }

    /// <summary>
    /// Health check endpoint to verify that the webhook receiver is reachable.
    /// </summary>
    [HttpGet("health")]
    public IActionResult Health() => Ok(new { status = "healthy", timestamp = DateTime.UtcNow });
}
