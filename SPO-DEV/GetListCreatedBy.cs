using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

class GetListCreatedBy
{
    static async Task Main(string[] args)
    {
        var clientId    = "YOUR_CLIENT_ID";
        var tenantId    = "YOUR_TENANT_ID";
        var siteUrl     = "https://contoso.sharepoint.com/sites/MyProject";
        var scopes      = new[] { $"{new Uri(siteUrl).Scheme}://{new Uri(siteUrl).Host}/AllSites.Manage" };

        var app = PublicClientApplicationBuilder
            .Create(clientId)
            .WithAuthority($"https://login.microsoftonline.com/{tenantId}")
            .WithRedirectUri("http://localhost")
            .Build();

        AuthenticationResult authResult;
        try
        {
            var accounts = await app.GetAccountsAsync();
            authResult = await app.AcquireTokenSilent(scopes, accounts.FirstOrDefault())
                                  .ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            authResult = await app.AcquireTokenInteractive(scopes)
                                  .ExecuteAsync();
        }

        using var ctx = new ClientContext(siteUrl);
        ctx.ExecutingWebRequest += (sender, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authResult.AccessToken;
        };

        var list = ctx.Web.Lists.GetByTitle("My Custom List");

        ctx.Load(list, l => l.Title,
                       l => l.Created,
                       l => l.Author.Title,
                       l => l.Author.Email,
                       l => l.Author.LoginName);

        ctx.ExecuteQuery();

        Console.WriteLine($"List:       {list.Title}");
        Console.WriteLine($"Created On: {list.Created}");
        Console.WriteLine($"Created By: {list.Author.Title}");
        Console.WriteLine($"Email:      {list.Author.Email}");
        Console.WriteLine($"Login:      {list.Author.LoginName}");
    }
}