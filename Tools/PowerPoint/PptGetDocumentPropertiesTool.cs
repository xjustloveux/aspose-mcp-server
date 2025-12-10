using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetDocumentPropertiesTool : IAsposeTool
{
    public string Description => "Get document properties (metadata)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var presentation = new Presentation(path);
        var props = presentation.DocumentProperties;
        var sb = new StringBuilder();

        sb.AppendLine("=== Document Properties ===");
        sb.AppendLine($"Title: {props.Title ?? "(none)"}");
        sb.AppendLine($"Subject: {props.Subject ?? "(none)"}");
        sb.AppendLine($"Author: {props.Author ?? "(none)"}");
        sb.AppendLine($"Keywords: {props.Keywords ?? "(none)"}");
        sb.AppendLine($"Comments: {props.Comments ?? "(none)"}");
        sb.AppendLine($"Category: {props.Category ?? "(none)"}");
        sb.AppendLine($"Company: {props.Company ?? "(none)"}");
        sb.AppendLine($"Manager: {props.Manager ?? "(none)"}");
        sb.AppendLine($"Created: {props.CreatedTime}");
        sb.AppendLine($"Modified: {props.LastSavedTime}");
        sb.AppendLine($"Revision: {props.RevisionNumber}");

        return await Task.FromResult(sb.ToString());
    }
}
