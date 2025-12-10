using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetParagraphsTool : IAsposeTool
{
    public string Description => "Get all paragraphs from Word document with optional filtering";

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
                description = "Section index (0-based, optional, if not provided returns all paragraphs)"
            },
            includeEmpty = new
            {
                type = "boolean",
                description = "Include empty paragraphs (optional, default: true)"
            },
            styleFilter = new
            {
                type = "string",
                description = "Filter by style name (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var includeEmpty = arguments?["includeEmpty"]?.GetValue<bool?>() ?? true;
        var styleFilter = arguments?["styleFilter"]?.GetValue<string>();

        var doc = new Document(path);
        var sb = new StringBuilder();

        List<Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            paragraphs = doc.Sections[sectionIndex.Value].Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        }
        else
        {
            paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        }

        if (!includeEmpty)
        {
            paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();
        }

        if (!string.IsNullOrEmpty(styleFilter))
        {
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();
        }

        sb.AppendLine($"=== Paragraphs ({paragraphs.Count}) ===");
        sb.AppendLine();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            sb.AppendLine($"[{i}] Style: {para.ParagraphFormat.Style?.Name ?? "(none)"}");
            sb.AppendLine($"    Text: {text.Substring(0, Math.Min(100, text.Length))}{(text.Length > 100 ? "..." : "")}");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

