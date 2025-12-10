using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordGetEndnotesTool : IAsposeTool
{
    public string Description => "Get all endnotes from Word document";

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

        sb.AppendLine("=== Endnotes ===");
        sb.AppendLine();

        var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();

        for (int i = 0; i < endnotes.Count; i++)
        {
            var endnote = endnotes[i];
            sb.AppendLine($"[{i + 1}] Reference Mark: {endnote.ReferenceMark}");
            sb.AppendLine($"    Text: {endnote.ToString(SaveFormat.Text).Trim()}");
            sb.AppendLine($"    Type: {endnote.FootnoteType}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Endnotes: {endnotes.Count}");

        return await Task.FromResult(sb.ToString());
    }
}

