using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetRevisionsTool : IAsposeTool
{
    public string Description => "Get all revisions from Word document";

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
        var sb = new StringBuilder();

        sb.AppendLine("=== Revisions ===");
        sb.AppendLine();

        var revisions = doc.Revisions.ToList();
        for (int i = 0; i < revisions.Count; i++)
        {
            var revision = revisions[i];
            sb.AppendLine($"[{i + 1}] Type: {revision.RevisionType}");
            sb.AppendLine($"    Author: {revision.Author}");
            sb.AppendLine($"    Date: {revision.DateTime}");
            sb.AppendLine($"    Text: {revision.ParentNode?.ToString(SaveFormat.Text)?.Trim() ?? "(none)"}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Revisions: {revisions.Count}");

        return await Task.FromResult(sb.ToString());
    }
}

