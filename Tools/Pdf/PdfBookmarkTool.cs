using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing bookmarks in PDF documents (add, delete, edit, get)
/// </summary>
public class PdfBookmarkTool : IAsposeTool
{
    public string Description => @"Manage bookmarks in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add bookmark: pdf_bookmark(operation='add', path='doc.pdf', title='Chapter 1', pageIndex=1)
- Delete bookmark: pdf_bookmark(operation='delete', path='doc.pdf', bookmarkIndex=0)
- Edit bookmark: pdf_bookmark(operation='edit', path='doc.pdf', bookmarkIndex=0, title='Updated Title', pageIndex=2)
- Get bookmarks: pdf_bookmark(operation='get', path='doc.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a bookmark (required params: path, title, pageIndex)
- 'delete': Delete a bookmark (required params: path, bookmarkIndex)
- 'edit': Edit a bookmark (required params: path, bookmarkIndex, title, pageIndex)
- 'get': Get all bookmarks (required params: path)",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
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
                description = "Bookmark index (1-based, required for delete, edit, optional for get)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "add" => await AddBookmark(arguments),
            "delete" => await DeleteBookmark(arguments),
            "edit" => await EditBookmark(arguments),
            "get" => await GetBookmarks(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a bookmark to the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, title, pageIndex, optional parentIndex, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> AddBookmark(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var title = ArgumentHelper.GetString(arguments, "title");
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

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
            return
                $"Successfully added bookmark '{title}' pointing to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a bookmark from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, bookmarkIndex, optional outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteBookmark(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var bookmarkIndex = ArgumentHelper.GetInt(arguments, "bookmarkIndex");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var document = new Document(path);
            if (bookmarkIndex < 1 || bookmarkIndex > document.Outlines.Count)
                throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

            var bookmark = document.Outlines[bookmarkIndex];
            var title = bookmark.Title;
            document.Outlines.Delete(title);
            document.Save(outputPath);
            return
                $"Successfully deleted bookmark '{title}' (index {bookmarkIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a bookmark
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, bookmarkIndex, optional title, pageIndex, outputPath</param>
    /// <returns>Success message</returns>
    private Task<string> EditBookmark(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var bookmarkIndex = ArgumentHelper.GetInt(arguments, "bookmarkIndex");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

            SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
            SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

            using var document = new Document(path);
            if (bookmarkIndex < 1 || bookmarkIndex > document.Outlines.Count)
                throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

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
            return $"Successfully edited bookmark (index {bookmarkIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all bookmarks from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with all bookmarks</returns>
    private Task<string> GetBookmarks(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);

            using var document = new Document(path);
            var sb = new StringBuilder();
            sb.AppendLine("=== PDF Bookmarks ===");
            sb.AppendLine();

            if (document.Outlines.Count == 0)
            {
                sb.AppendLine("No bookmarks found.");
                return sb.ToString();
            }

            sb.AppendLine($"Total Bookmarks: {document.Outlines.Count}");
            sb.AppendLine();

            for (var i = 1; i <= document.Outlines.Count; i++)
            {
                var bookmark = document.Outlines[i];
                sb.AppendLine($"[{i}] Title: {bookmark.Title}");
                if (bookmark.Action is GoToAction { Destination: XYZExplicitDestination xyzDest })
                {
                    var pageNum = document.Pages.IndexOf(xyzDest.Page) + 1;
                    sb.AppendLine($"    Page: {pageNum}");
                }

                sb.AppendLine();
            }

            return sb.ToString();
        });
    }
}