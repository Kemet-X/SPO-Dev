# SharePoint Online Remote Event Receiver Azure Function

This project is an Azure Function that acts as a Remote Event Receiver for SharePoint Online. It supports both certificate-based and client secret authentication.

## Project Structure
- **SharePoint-RER-AzureFunction**
  - **Function.cs**: Main entry point for the Azure Function.
  - **local.settings.json**: Contains local settings for the Azure Function.
  - **function.json**: Configuration for the Azure Function.
  - **README.md**: Documentation for setting up and running the Azure Function.
  - **script.sh**: Any necessary scripts for deployment.

## Authentication Methods
- **Certificate-based Authentication**: Utilizes a certificate for authenticating to SharePoint Online.
- **Client Secret Authentication**: Uses client ID and client secret for authentication.

## Setup Instructions
1. Clone the repository.
2. Install required dependencies.
3. Set up Azure Function in Azure portal.
4. Configure authentication settings in Azure portal.
5. Deploy the function.

## Examples
```csharp
using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;

public static class Function
{
    [FunctionName("SharePointEventReceiver")]    
    public static void Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        log.LogInformation("C# HTTP trigger function processed a request.");
        // Implement your logic here
    }
}
```
