using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

class CreateListWithCSOM
{
    static async Task Main(string[] args)
    {
        var clientId    = "YOUR_CLIENT_ID";
        var tenantId    = "YOUR_TENANT_ID";
        var siteUrl     = "https://contoso.sharepoint.com/sites/MyProject";

        var scopes = new[] { $"{new Uri(siteUrl).Scheme}://{new Uri(siteUrl).Host}/AllSites.Manage" };

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

        Console.WriteLine($"Signed in as: {authResult.Account.Username}");

        using var ctx = new ClientContext(siteUrl);
        ctx.ExecutingWebRequest += (sender, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authResult.AccessToken;
        };

        var listCreationInfo = new ListCreationInformation
        {
            Title = "My Custom List",
            TemplateType = (int)ListTemplateType.GenericList
        };

        var list = ctx.Web.Lists.Add(listCreationInfo);

        list.Fields.AddFieldAsXml(
            @"<Field DisplayName='Project Name' Name='ProjectName' Type='Text' />",
            true, AddFieldOptions.DefaultValue);

        list.Fields.AddFieldAsXml(
            @"<Field DisplayName='Status' Name='Status' Type='Choice'>
                <CHOICES>
                    <CHOICE>Not Started</CHOICE>
                    <CHOICE>In Progress</CHOICE>
                    <CHOICE>Completed</CHOICE>
                </CHOICES>
              </Field>",
            true, AddFieldOptions.DefaultValue);

        list.Fields.AddFieldAsXml(
            @"<Field DisplayName='Due Date' Name='DueDate' Type='DateTime' Format='DateOnly' />",
            true, AddFieldOptions.DefaultValue);

        ctx.ExecuteQuery();
        Console.WriteLine("List created successfully as the signed-in user!");
    }
}