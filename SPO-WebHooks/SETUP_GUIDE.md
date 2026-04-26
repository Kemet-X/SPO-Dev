# Setup Guide for SharePoint Online Webhooks with Azure AD Certificate Authentication

## Introduction
This guide will help you set up SharePoint Online webhooks using Azure AD certificate authentication. This is essential for securely receiving notifications from SharePoint when content changes are made.

## Prerequisites
- An Azure subscription
- Access to SharePoint Online
- Azure Active Directory (AD) admin rights
- A certificate for authentication (self-signed or from a CA)

## Step 1: Create and Upload a Certificate
1. Generate a self-signed certificate if you don’t have one. You can use PowerShell or OpenSSL for this.
   
   ### PowerShell Command Example:
   ```powershell
   New-SelfSignedCertificate -DnsName "YourAppName" -CertStoreLocation "Cert:\\CurrentUser\Personal"
   ```

2. Navigate to Azure AD admin center.
3. Go to **App registrations** -> **Your App**.
4. Under **Certificates & secrets**, upload your certificate. Make sure to note the thumbprint.

## Step 2: Register Your Application in Azure AD
1. In the Azure AD portal, register a new application.
2. Set the redirect URI. Applications that listen for SharePoint webhooks need to have a valid HTTPS redirect URI.
3. Grant the application API permissions for SharePoint. Choose permissions like `Sites.Read.All` or others depending on your needs.
4. Note down the Application (client) ID and Directory (tenant) ID for later use.

## Step 3: Enable Webhooks on SharePoint List
1. Go to the SharePoint site and navigate to the list where you want to enable webhooks.
2. Use the following HTTP POST request to create a webhook. Make sure you use your app's client ID and the URL of your webhook listener:
   
   ```http
   POST https://<your-tenant>.sharepoint.com/sites/<your-site>/_api/Web/Lists/getByTitle('<your-list-name>')/Subscriptions
   Content-Type: application/json
   Authorization: Bearer <your_access_token>
   
   {
       "resource": "https://<your-tenant>.sharepoint.com/sites/<your-site>/_api/Web/Lists/getByTitle('<your-list-name>')",
       "notificationUrl": "https://<your-notification-url>",
       "expirationDateTime": "2026-05-26T20:42:55Z",
       "clientState": "secret-client-value"
   }
   ```
3. You will receive a validation token at your endpoint. You must return it to validate the webhook subscription.

## Step 4: Handling Incoming Notifications
1. Ensure your notification endpoint can parse the incoming requests.
2. Implement logic to handle the different types of notifications received from SharePoint.

## Conclusion
You have now set up SharePoint Online webhooks with Azure AD certificate authentication. Make sure to monitor your application for any authentication or webhook delivery issues.

## Additional Resources
- [Microsoft Documentation on Webhooks](https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/register-sharepoint-add-ins)
- [Azure AD Certificates](https://docs.microsoft.com/en-us/azure/active-directory/develop/registered-apps)
