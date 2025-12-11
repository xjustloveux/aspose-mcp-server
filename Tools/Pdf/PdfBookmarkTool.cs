using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfBookmarkTool : IAsposeTool
{
    public string Description => "Manage bookmarks in PDF documents (add, delete, edit, get)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add, delete, edit, get",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            title = new
            {
                type = "string",
                description = "Bookmark title (required for add, edit)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Target page index (1-based, required for add, edit)"
            },
            bookmarkIndex = new
            {
                type = "number",
                description = "Bookmark index (0-based, required for delete, edit)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddBookmark(arguments),
            "delete" => await DeleteBookmark(arguments),
            "edit" => await EditBookmark(arguments),
            "get" => await GetBookmarks(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddBookmark(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var title = arguments?["title"]?.GetValue<string>() ?? throw new ArgumentException("title is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var bookmark = new OutlineItemCollection(document.Outlines)
        {
            Title = title,
            Action = new GoToAction(document.Pages[pageIndex])
        };

        document.Outlines.Add(bookmark);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added bookmark '{title}' pointing to page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> DeleteBookmark(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var bookmarkIndex = arguments?["bookmarkIndex"]?.GetValue<int>() ?? throw new ArgumentException("bookmarkIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (bookmarkIndex < 0 || bookmarkIndex >= document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 0 and {document.Outlines.Count - 1}");

        var bookmark = document.Outlines[bookmarkIndex];
        var title = bookmark.Title;
        document.Outlines.Delete(title);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted bookmark '{title}' (index {bookmarkIndex}). Output: {outputPath}");
    }

    private async Task<string> EditBookmark(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var bookmarkIndex = arguments?["bookmarkIndex"]?.GetValue<int>() ?? throw new ArgumentException("bookmarkIndex is required");
        var title = arguments?["title"]?.GetValue<string>();
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (bookmarkIndex < 0 || bookmarkIndex >= document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 0 and {document.Outlines.Count - 1}");

        var bookmark = document.Outlines[bookmarkIndex];
        
        if (!string.IsNullOrEmpty(title))
            bookmark.Title = title;
        
        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            bookmark.Action = new GoToAction(document.Pages[pageIndex.Value]);
        }

        document.Save(outputPath);
        return await Task.FromResult($"Successfully edited bookmark (index {bookmarkIndex}). Output: {outputPath}");
    }

    private async Task<string> GetBookmarks(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Bookmarks ===");
        sb.AppendLine();

        if (document.Outlines.Count == 0)
        {
            sb.AppendLine("No bookmarks found.");
            return await Task.FromResult(sb.ToString());
        }

        sb.AppendLine($"Total Bookmarks: {document.Outlines.Count}");
        sb.AppendLine();

        for (int i = 0; i < document.Outlines.Count; i++)
        {
            var bookmark = document.Outlines[i];
            sb.AppendLine($"[{i}] Title: {bookmark.Title}");
            if (bookmark.Action is GoToAction goToAction && goToAction.Destination is Aspose.Pdf.Annotations.XYZExplicitDestination xyzDest)
            {
                var pageNum = document.Pages.IndexOf(xyzDest.Page) + 1;
                sb.AppendLine($"    Page: {pageNum}");
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

