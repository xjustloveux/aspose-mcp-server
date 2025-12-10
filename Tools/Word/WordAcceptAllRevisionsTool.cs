using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAcceptAllRevisionsTool : IAsposeTool
{
    public string Description => "Accept all revisions in Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        var doc = new Document(path);
        doc.AcceptAllRevisions();

        doc.Save(path);
        return await Task.FromResult($"All revisions accepted: {path}");
    }
}

