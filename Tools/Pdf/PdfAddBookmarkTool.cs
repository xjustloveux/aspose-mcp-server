using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfAddBookmarkTool : IAsposeTool
{
    public string Description => "Add a bookmark to a PDF document";

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
            title = new
            {
                type = "string",
                description = "Bookmark title"
            },
            pageIndex = new
            {
                type = "number",
                description = "Target page index (1-based)"
            }
        },
        required = new[] { "path", "title", "pageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var title = arguments?["title"]?.GetValue<string>() ?? throw new ArgumentException("title is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        using var document = new Document(path);
        
        var bookmark = new OutlineItemCollection(document.Outlines)
        {
            Title = title,
            Action = new GoToAction(document.Pages[pageIndex])
        };

        document.Outlines.Add(bookmark);
        document.Save(path);

        return await Task.FromResult($"Bookmark '{title}' added to PDF: {path}");
    }
}

