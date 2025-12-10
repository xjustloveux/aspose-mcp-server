using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Properties;

namespace AsposeMcpServer.Tools;

public class WordGetDocumentPropertiesTool : IAsposeTool
{
    public string Description => "Get document properties (metadata) from Word document";

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
        var props = doc.BuiltInDocumentProperties;
        var customProps = doc.CustomDocumentProperties;
        var sb = new StringBuilder();

        sb.AppendLine("=== Document Properties ===");
        sb.AppendLine();
        sb.AppendLine("Built-in Properties:");
        sb.AppendLine($"  Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"  Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"  Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"  Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"  Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"  Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"  Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"  Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"  Created: {props.CreatedTime}");
        sb.AppendLine($"  Modified: {props.LastSavedTime}");
        sb.AppendLine($"  Last Saved By: {props.LastSavedBy ?? "(none)"}");
        sb.AppendLine($"  Revision: {props.RevisionNumber}");
        sb.AppendLine($"  Word Count: {props.Words}");
        sb.AppendLine($"  Character Count: {props.Characters}");
        sb.AppendLine($"  Page Count: {props.Pages}");

        if (customProps.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Custom Properties:");
            foreach (DocumentProperty prop in customProps)
            {
                sb.AppendLine($"  {prop.Name}: {prop.Value}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

