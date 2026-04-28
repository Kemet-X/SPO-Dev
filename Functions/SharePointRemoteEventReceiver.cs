using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

[Function("SharePointRemoteEventReceiver")]
public async Task<HttpResponseData> RunAsync(
    [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequestData req,
    ILogger log,
    FunctionContext executionContext)
{
    log.LogInformation("SharePoint Remote Event Receiver triggered");
    
    try
    {
        // Read request body
        using var streamReader = new StreamReader(req.Body);
        var requestBody = await streamReader.ReadToEndAsync();
        log.LogInformation($"Request body: {requestBody}");
        
        // Parse the event
        using var jsonDoc = JsonDocument.Parse(requestBody);
        var root = jsonDoc.RootElement;
        
        var eventType = root.GetProperty("eventType").GetString();
        log.LogInformation($"Event type: {eventType}");
        
        // Route to appropriate handler
        SharePointEventResponse response;
        
        if (eventType == "ItemAdding")
        {
            var handler = executionContext.InstanceServices.GetService(typeof(ItemAddingEventHandler)) as ItemAddingEventHandler;
            response = await handler.HandleItemAdding(root);
        }
        else
        {
            response = new SharePointEventResponse
            {
                ErrorMessage = "Event type not supported",
                Status = "Continue"
            };
        }
        
        // Return response
        var httpResponse = req.CreateResponse(System.Net.HttpStatusCode.OK);
        await httpResponse.WriteAsJsonAsync(response);
        
        return httpResponse;
    }
    catch (Exception ex)
    {
        log.LogError($"Error processing event: {ex.Message}");
        
        var errorResponse = req.CreateResponse(System.Net.HttpStatusCode.InternalServerError);
        await errorResponse.WriteAsJsonAsync(new { error = ex.Message });
        
        return errorResponse;
    }
}

/// <summary>
/// Handles ItemAdding events from SharePoint
/// </summary>
public class ItemAddingEventHandler
{
    private readonly TitleValidationService _validationService;
    private readonly ILogger<ItemAddingEventHandler> _logger;
    
    public ItemAddingEventHandler(TitleValidationService validationService, ILogger<ItemAddingEventHandler> logger)
    {
        _validationService = validationService;
        _logger = logger;
    }
    
    public async Task<SharePointEventResponse> HandleItemAdding(JsonElement eventData)
    {
        try
        {
            _logger.LogInformation("Processing ItemAdding event");
            
            // Extract AfterProperties
            if (!eventData.TryGetProperty("itemEventProperties", out var eventProps))
            {
                return new SharePointEventResponse
                {
                    Status = "Continue",
                    ErrorMessage = null
                };
            }
            
            if (!eventProps.TryGetProperty("AfterProperties", out var afterProps))
            {
                return new SharePointEventResponse
                {
                    Status = "Continue",
                    ErrorMessage = null
                };
            }
            
            // Get Title
            string title = null;
            if (afterProps.TryGetProperty("Title", out var titleElement))
            {
                title = titleElement.GetString();
            }
            
            _logger.LogInformation($"Title: {title}");
            
            // Validate title
            if (!string.IsNullOrEmpty(title))
            {
                var validationResult = _validationService.ValidateTitle(title);
                
                if (!validationResult.IsValid)
                {
                    _logger.LogWarning($"Title validation failed: {validationResult.ErrorMessage}");
                    
                    return new SharePointEventResponse
                    {
                        Status = "CancelWithError",
                        ErrorMessage = validationResult.ErrorMessage
                    };
                }
            }
            
            _logger.LogInformation("Item allowed - title is valid");
            
            return new SharePointEventResponse
            {
                Status = "Continue",
                ErrorMessage = null
            };
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error handling ItemAdding event: {ex.Message}");
            
            return new SharePointEventResponse
            {
                Status = "Continue",
                ErrorMessage = null
            };
        }
    }
}

/// <summary>
/// Validates SharePoint list item titles
/// </summary>
public class TitleValidationService
{
    private readonly ILogger<TitleValidationService> _logger;
    
    public TitleValidationService(ILogger<TitleValidationService> logger)
    {
        _logger = logger;
    }
    
    public ValidationResult ValidateTitle(string title)
    {
        _logger.LogInformation($"Validating title: {title}");
        
        // Check if title contains any digits
        if (title.Any(char.IsDigit))
        {
            return new ValidationResult
            {
                IsValid = false,
                ErrorMessage = "❌ Item creation cancelled: Title cannot contain numbers. Please enter a title with letters and symbols only."
            };
        }
        
        return new ValidationResult { IsValid = true };
    }
}

public class ValidationResult
{
    public bool IsValid { get; set; }
    public string ErrorMessage { get; set; }
}

public class SharePointEventResponse
{
    [System.Text.Json.Serialization.JsonPropertyName("status")]
    public string Status { get; set; }
    
    [System.Text.Json.Serialization.JsonPropertyName("errorMessage")]
    public string ErrorMessage { get; set; }
}