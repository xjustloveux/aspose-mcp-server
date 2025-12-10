using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordGetFootnotesTool : IAsposeTool
{
    public string Description => "Get all footnotes from Word document";

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

        sb.AppendLine("=== Footnotes ===");
        sb.AppendLine();

        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Footnote)
            .ToList();

        for (int i = 0; i < footnotes.Count; i++)
        {
            var footnote = footnotes[i];
            sb.AppendLine($"[{i + 1}] Reference Mark: {footnote.ReferenceMark}");
            sb.AppendLine($"    Text: {footnote.ToString(SaveFormat.Text).Trim()}");
            sb.AppendLine($"    Type: {footnote.FootnoteType}");
            sb.AppendLine();
        }

        sb.AppendLine($"Total Footnotes: {footnotes.Count}");

        return await Task.FromResult(sb.ToString());
    }
}

