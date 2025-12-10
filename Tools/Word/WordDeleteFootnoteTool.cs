using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordDeleteFootnoteTool : IAsposeTool
{
    public string Description => "Delete footnote(s) from Word document";

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
            referenceMark = new
            {
                type = "string",
                description = "Reference mark of footnote to delete (optional, if not provided deletes all footnotes)"
            },
            footnoteIndex = new
            {
                type = "number",
                description = "Footnote index (0-based, optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var footnoteIndex = arguments?["footnoteIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Footnote)
            .ToList();

        int deletedCount = 0;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            var footnote = footnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            if (footnote != null)
            {
                footnote.Remove();
                deletedCount = 1;
            }
        }
        else if (footnoteIndex.HasValue)
        {
            if (footnoteIndex.Value >= 0 && footnoteIndex.Value < footnotes.Count)
            {
                footnotes[footnoteIndex.Value].Remove();
                deletedCount = 1;
            }
        }
        else
        {
            // Delete all footnotes
            foreach (var footnote in footnotes)
            {
                footnote.Remove();
                deletedCount++;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Deleted {deletedCount} footnote(s): {path}");
    }
}

