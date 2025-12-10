using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetSectionsInfoTool : IAsposeTool
{
    public string Description => "Get detailed information about all sections in Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided returns all sections)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var sb = new StringBuilder();

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }

            var section = doc.Sections[sectionIndex.Value];
            sb.AppendLine($"=== Section {sectionIndex.Value} ===");
            sb.AppendLine($"  Page Setup:");
            sb.AppendLine($"    Page Width: {section.PageSetup.PageWidth}");
            sb.AppendLine($"    Page Height: {section.PageSetup.PageHeight}");
            sb.AppendLine($"    Orientation: {section.PageSetup.Orientation}");
            sb.AppendLine($"    Margins: Left={section.PageSetup.LeftMargin}, Right={section.PageSetup.RightMargin}, Top={section.PageSetup.TopMargin}, Bottom={section.PageSetup.BottomMargin}");
            sb.AppendLine($"  Paragraphs: {section.Body.GetChildNodes(NodeType.Paragraph, true).Count}");
            sb.AppendLine($"  Tables: {section.Body.GetChildNodes(NodeType.Table, true).Count}");
            sb.AppendLine($"  Headers/Footers: {section.HeadersFooters.Count}");
        }
        else
        {
            sb.AppendLine($"=== Sections ({doc.Sections.Count}) ===");
            sb.AppendLine();

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                var section = doc.Sections[i];
                sb.AppendLine($"Section {i}:");
                sb.AppendLine($"  Page Setup: {section.PageSetup.PageWidth}x{section.PageSetup.PageHeight}, {section.PageSetup.Orientation}");
                sb.AppendLine($"  Paragraphs: {section.Body.GetChildNodes(NodeType.Paragraph, true).Count}");
                sb.AppendLine($"  Tables: {section.Body.GetChildNodes(NodeType.Table, true).Count}");
                sb.AppendLine();
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

