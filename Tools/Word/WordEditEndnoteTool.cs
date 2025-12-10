using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordEditEndnoteTool : IAsposeTool
{
    public string Description => "Edit endnote text in Word document";

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
                description = "Reference mark of endnote to edit (optional)"
            },
            endnoteIndex = new
            {
                type = "number",
                description = "Endnote index (0-based, optional)"
            },
            newText = new
            {
                type = "string",
                description = "New endnote text"
            }
        },
        required = new[] { "path", "newText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var endnoteIndex = arguments?["endnoteIndex"]?.GetValue<int?>();
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");

        var doc = new Document(path);
        var endnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Endnote)
            .ToList();

        Footnote? endnote = null;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            endnote = endnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
        }
        else if (endnoteIndex.HasValue)
        {
            if (endnoteIndex.Value >= 0 && endnoteIndex.Value < endnotes.Count)
            {
                endnote = endnotes[endnoteIndex.Value];
            }
        }
        else if (endnotes.Count > 0)
        {
            endnote = endnotes[0];
        }

        if (endnote == null)
        {
            throw new ArgumentException("Endnote not found");
        }

        endnote.RemoveAllChildren();
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(endnote.FirstParagraph);
        builder.Write(newText);

        doc.Save(path);
        return await Task.FromResult($"Endnote edited: {path}");
    }
}

