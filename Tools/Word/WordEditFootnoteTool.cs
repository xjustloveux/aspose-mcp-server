using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Notes;

namespace AsposeMcpServer.Tools;

public class WordEditFootnoteTool : IAsposeTool
{
    public string Description => "Edit footnote text in Word document";

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
                description = "Reference mark of footnote to edit (optional)"
            },
            footnoteIndex = new
            {
                type = "number",
                description = "Footnote index (0-based, optional)"
            },
            newText = new
            {
                type = "string",
                description = "New footnote text"
            }
        },
        required = new[] { "path", "newText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var referenceMark = arguments?["referenceMark"]?.GetValue<string>();
        var footnoteIndex = arguments?["footnoteIndex"]?.GetValue<int?>();
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");

        var doc = new Document(path);
        var footnotes = doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == FootnoteType.Footnote)
            .ToList();

        Footnote? footnote = null;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            footnote = footnotes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
        }
        else if (footnoteIndex.HasValue)
        {
            if (footnoteIndex.Value >= 0 && footnoteIndex.Value < footnotes.Count)
            {
                footnote = footnotes[footnoteIndex.Value];
            }
        }
        else if (footnotes.Count > 0)
        {
            footnote = footnotes[0];
        }

        if (footnote == null)
        {
            throw new ArgumentException("Footnote not found");
        }

        footnote.RemoveAllChildren();
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(footnote.FirstParagraph);
        builder.Write(newText);

        doc.Save(path);
        return await Task.FromResult($"Footnote edited: {path}");
    }
}

