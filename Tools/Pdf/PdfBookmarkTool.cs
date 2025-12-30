using System.Text.Json;
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
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

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
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddBookmark(path, outputPath!, arguments),
            "delete" => await DeleteBookmark(path, outputPath!, arguments),
            "edit" => await EditBookmark(path, outputPath!, arguments),
            "get" => await GetBookmarks(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a bookmark to the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing title, pageIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range</exception>
    private Task<string> AddBookmark(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var title = ArgumentHelper.GetString(arguments, "title");
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

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
            return $"Added bookmark '{title}' pointing to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a bookmark from the PDF by its index
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing bookmarkIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when bookmarkIndex is out of range</exception>
    /// <remarks>Note: Deletes by title, which may affect multiple bookmarks with the same title</remarks>
    private Task<string> DeleteBookmark(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var bookmarkIndex = ArgumentHelper.GetInt(arguments, "bookmarkIndex");

            using var document = new Document(path);
            if (bookmarkIndex < 1 || bookmarkIndex > document.Outlines.Count)
                throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

            var bookmark = document.Outlines[bookmarkIndex];
            var title = bookmark.Title;

            var countBefore = document.Outlines.Count;
            document.Outlines.Delete(title);
            var countAfter = document.Outlines.Count;
            var deletedCount = countBefore - countAfter;

            document.Save(outputPath);

            return deletedCount > 1
                ? $"Deleted {deletedCount} bookmark(s) with title '{title}' (requested index {bookmarkIndex}). Output: {outputPath}"
                : $"Deleted bookmark '{title}' (index {bookmarkIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a bookmark
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing bookmarkIndex, optional title, pageIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when bookmarkIndex or pageIndex is out of range</exception>
    private Task<string> EditBookmark(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var bookmarkIndex = ArgumentHelper.GetInt(arguments, "bookmarkIndex");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

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
            return $"Edited bookmark (index {bookmarkIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all bookmarks from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with all bookmarks</returns>
    private Task<string> GetBookmarks(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);

            if (document.Outlines.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No bookmarks found"
                };
                return JsonSerializer.Serialize(emptyResult, JsonOptions);
            }

            var bookmarkList = new List<object>();
            var index = 0;

            foreach (var bookmark in document.Outlines)
            {
                index++;
                var bookmarkInfo = new Dictionary<string, object?>
                {
                    ["index"] = index,
                    ["title"] = bookmark.Title
                };

                var pageIndex = ExtractPageIndex(bookmark, document);
                if (pageIndex.HasValue)
                    bookmarkInfo["pageIndex"] = pageIndex.Value;

                bookmarkList.Add(bookmarkInfo);
            }

            var result = new
            {
                count = bookmarkList.Count,
                items = bookmarkList
            };
            return JsonSerializer.Serialize(result, JsonOptions);
        });
    }

    /// <summary>
    ///     Extracts the page index from a bookmark's destination or action
    /// </summary>
    /// <param name="bookmark">The bookmark to extract page index from</param>
    /// <param name="document">The PDF document</param>
    /// <returns>The 1-based page index, or null if not found</returns>
    private static int? ExtractPageIndex(OutlineItemCollection bookmark, Document document)
    {
        Page? page = null;

        // Try to get page from Destination property
        if (bookmark.Destination is ExplicitDestination explicitDest)
            page = explicitDest.Page;
        // Try to get page from GoToAction
        else if (bookmark.Action is GoToAction { Destination: ExplicitDestination actionDest })
            page = actionDest.Page;

        if (page == null)
            return null;

        var pageIndex = document.Pages.IndexOf(page);
        return pageIndex > 0 ? pageIndex : null;
    }
}