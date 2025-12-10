using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordDeleteEndnoteTool : IAsposeTool
{
    public string Description => "Delete endnote(s) from Word document";

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
                description = "Reference mark of endnote to delete (optional, if not provided deletes all endnotes)"
            },
            endnoteIndex = new
            {
                type = "number",
                description = "Endnote index (0-based, optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var endnoteIndex = arguments?["endnoteIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();

        int deletedCount = 0;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            var endnote = endnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            if (endnote != null)
            {
                endnote.Remove();
                deletedCount = 1;
            }
        }
        else if (endnoteIndex.HasValue)
        {
            if (endnoteIndex.Value >= 0 && endnoteIndex.Value < endnotes.Count)
            {
                endnotes[endnoteIndex.Value].Remove();
                deletedCount = 1;
            }
        }
        else
        {
            // Delete all endnotes
            foreach (var endnote in endnotes)
            {
                endnote.Remove();
                deletedCount++;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Deleted {deletedCount} endnote(s): {path}");
    }
}

