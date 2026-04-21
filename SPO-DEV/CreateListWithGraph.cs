using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

class CreateListWithGraph
{
    static async Task Main(string[] args)
    {
        var tenantId     = "YOUR_TENANT_ID";
        var clientId     = "YOUR_CLIENT_ID";
        var clientSecret = "YOUR_CLIENT_SECRET";
        var siteId       = "YOUR_SITE_ID";

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        var graphClient = new GraphServiceClient(credential);

        var list = new List
        {
            DisplayName = "My Custom List",
            ListProp = new ListInfo
            {
                Template = "genericList"
            },
            Columns = new System.Collections.Generic.List<ColumnDefinition>
            {
                new ColumnDefinition
                {
                    Name = "ProjectName",
                    DisplayName = "Project Name",
                    Text = new TextColumn { MaxLength = 255 }
                },
                new ColumnDefinition
                {
                    Name = "Status",
                    DisplayName = "Status",
                    Choice = new ChoiceColumn
                    {
                        Choices = new System.Collections.Generic.List<string>
                            { "Not Started", "In Progress", "Completed" }
                    }
                },
                new ColumnDefinition
                {
                    Name = "DueDate",
                    DisplayName = "Due Date",
                    DateTime = new DateTimeColumn { Format = "dateOnly" }
                }
            }
        };

        try
        {
            var createdList = await graphClient.Sites[siteId].Lists.PostAsync(list);
            Console.WriteLine($"List created: {createdList?.DisplayName} (ID: {createdList?.Id})");
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            Console.WriteLine($"Error: {ex.Error?.Code} - {ex.Error?.Message}");
        }
    }
}