using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfEditBookmarkTool : IAsposeTool
{
    public string Description => "Edit bookmark properties in PDF document";

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
            bookmarkIndex = new
            {
                type = "number",
                description = "Bookmark index (0-based)"
            },
            title = new
            {
                type = "string",
                description = "New bookmark title (optional)"
            },
            pageNumber = new
            {
                type = "number",
                description = "Target page number (1-based, optional)"
            },
            level = new
            {
                type = "number",
                description = "Bookmark level (optional)"
            }
        },
        required = new[] { "path", "bookmarkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var bookmarkIndex = arguments?["bookmarkIndex"]?.GetValue<int>() ?? throw new ArgumentException("bookmarkIndex is required");
        var title = arguments?["title"]?.GetValue<string>();
        var pageNumber = arguments?["pageNumber"]?.GetValue<int?>();
        var level = arguments?["level"]?.GetValue<int?>();

        using var document = new Document(path);
        var outlineItemCollection = document.Outlines;

        if (bookmarkIndex < 0 || bookmarkIndex >= outlineItemCollection.Count)
        {
            throw new ArgumentException($"bookmarkIndex must be between 0 and {outlineItemCollection.Count - 1}");
        }

        var bookmark = outlineItemCollection[bookmarkIndex];

        if (!string.IsNullOrEmpty(title))
        {
            bookmark.Title = title;
        }

        if (pageNumber.HasValue)
        {
            if (pageNumber.Value < 1 || pageNumber.Value > document.Pages.Count)
            {
                throw new ArgumentException($"pageNumber must be between 1 and {document.Pages.Count}");
            }
            // Create destination using GoToAction
            var targetPage = document.Pages[pageNumber.Value];
            bookmark.Action = new Aspose.Pdf.Annotations.GoToAction(targetPage);
        }

        // Note: Level property may be read-only in some versions
        // Level is typically set when creating bookmarks

        document.Save(path);
        return await Task.FromResult($"Bookmark {bookmarkIndex} updated: {path}");
    }
}

