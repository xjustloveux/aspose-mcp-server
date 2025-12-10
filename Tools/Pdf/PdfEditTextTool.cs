using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfEditTextTool : IAsposeTool
{
    public string Description => "Edit text in PDF document (replace existing text fragment)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            oldText = new
            {
                type = "string",
                description = "Text to replace (searches for first occurrence)"
            },
            newText = new
            {
                type = "string",
                description = "New text"
            },
            replaceAll = new
            {
                type = "boolean",
                description = "Replace all occurrences (optional, default: false)"
            }
        },
        required = new[] { "path", "pageIndex", "oldText", "newText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var oldText = arguments?["oldText"]?.GetValue<string>() ?? throw new ArgumentException("oldText is required");
        var newText = arguments?["newText"]?.GetValue<string>() ?? throw new ArgumentException("newText is required");
        var replaceAll = arguments?["replaceAll"]?.GetValue<bool?>() ?? false;

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var textFragmentAbsorber = new TextFragmentAbsorber(oldText);
        page.Accept(textFragmentAbsorber);

        var count = 0;
        if (replaceAll)
        {
            foreach (TextFragment fragment in textFragmentAbsorber.TextFragments)
            {
                fragment.Text = newText;
                count++;
            }
        }
        else
        {
            if (textFragmentAbsorber.TextFragments.Count > 0)
            {
                textFragmentAbsorber.TextFragments[1].Text = newText;
                count = 1;
            }
        }

        document.Save(path);
        return await Task.FromResult($"Replaced {count} occurrence(s) of text on page {pageIndex}: {path}");
    }
}

