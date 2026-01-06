using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing bookmarks in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfBookmarkTool
{
    /// <summary>
    ///     JSON serialization options for formatted output.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfBookmarkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfBookmarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PDF bookmark operation (add, delete, edit, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, edit, get.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="title">Bookmark title (required for add, edit).</param>
    /// <param name="pageIndex">Target page index (1-based, required for add, edit).</param>
    /// <param name="bookmarkIndex">Bookmark index (1-based, required for delete, edit).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_bookmark")]
    [Description(@"Manage bookmarks in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add bookmark: pdf_bookmark(operation='add', path='doc.pdf', title='Chapter 1', pageIndex=1)
- Delete bookmark: pdf_bookmark(operation='delete', path='doc.pdf', bookmarkIndex=0)
- Edit bookmark: pdf_bookmark(operation='edit', path='doc.pdf', bookmarkIndex=0, title='Updated Title', pageIndex=2)
- Get bookmarks: pdf_bookmark(operation='get', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a bookmark (required params: path, title, pageIndex)
- 'delete': Delete a bookmark (required params: path, bookmarkIndex)
- 'edit': Edit a bookmark (required params: path, bookmarkIndex, title, pageIndex)
- 'get': Get all bookmarks (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Bookmark title (required for add, edit)")]
        string? title = null,
        [Description("Target page index (1-based, required for add, edit)")]
        int? pageIndex = null,
        [Description("Bookmark index (1-based, required for delete, edit, optional for get)")]
        int? bookmarkIndex = null)
    {
        return operation.ToLower() switch
        {
            "add" => AddBookmark(sessionId, path, outputPath, title, pageIndex),
            "delete" => DeleteBookmark(sessionId, path, outputPath, bookmarkIndex),
            "edit" => EditBookmark(sessionId, path, outputPath, bookmarkIndex, title, pageIndex),
            "get" => GetBookmarks(sessionId, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new bookmark to the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="title">The bookmark title.</param>
    /// <param name="pageIndex">The 1-based target page index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string AddBookmark(string? sessionId, string? path, string? outputPath, string? title, int? pageIndex)
    {
        if (string.IsNullOrEmpty(title))
            throw new ArgumentException("title is required for add operation");
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for add operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var bookmark = new OutlineItemCollection(document.Outlines)
        {
            Title = title,
            Action = new GoToAction(document.Pages[pageIndex.Value])
        };

        document.Outlines.Add(bookmark);
        ctx.Save(outputPath);
        return $"Added bookmark '{title}' pointing to page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a bookmark from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="bookmarkIndex">The 1-based bookmark index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the bookmark index is invalid.</exception>
    private string DeleteBookmark(string? sessionId, string? path, string? outputPath, int? bookmarkIndex)
    {
        if (!bookmarkIndex.HasValue)
            throw new ArgumentException("bookmarkIndex is required for delete operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (bookmarkIndex.Value < 1 || bookmarkIndex.Value > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[bookmarkIndex.Value];
        var title = bookmark.Title;

        var countBefore = document.Outlines.Count;
        document.Outlines.Delete(title);
        var countAfter = document.Outlines.Count;
        var deletedCount = countBefore - countAfter;

        ctx.Save(outputPath);

        return deletedCount > 1
            ? $"Deleted {deletedCount} bookmark(s) with title '{title}' (requested index {bookmarkIndex.Value}). {ctx.GetOutputMessage(outputPath)}"
            : $"Deleted bookmark '{title}' (index {bookmarkIndex.Value}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing bookmark in the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="bookmarkIndex">The 1-based bookmark index.</param>
    /// <param name="title">The new bookmark title.</param>
    /// <param name="pageIndex">The new 1-based target page index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the bookmark index is invalid.</exception>
    private string EditBookmark(string? sessionId, string? path, string? outputPath, int? bookmarkIndex, string? title,
        int? pageIndex)
    {
        if (!bookmarkIndex.HasValue)
            throw new ArgumentException("bookmarkIndex is required for edit operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

        if (bookmarkIndex.Value < 1 || bookmarkIndex.Value > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[bookmarkIndex.Value];

        if (!string.IsNullOrEmpty(title))
            bookmark.Title = title;

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            bookmark.Action = new GoToAction(document.Pages[pageIndex.Value]);
        }

        ctx.Save(outputPath);
        return $"Edited bookmark (index {bookmarkIndex.Value}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves all bookmarks from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <returns>A JSON string containing bookmark information.</returns>
    private string GetBookmarks(string? sessionId, string? path)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var document = ctx.Document;

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

        List<object> bookmarkList = [];
        var index = 0;

        foreach (var bookmark in document.Outlines)
        {
            index++;
            var bookmarkInfo = new Dictionary<string, object?>
            {
                ["index"] = index,
                ["title"] = bookmark.Title
            };

            var extractedPageIndex = ExtractPageIndex(bookmark, document);
            if (extractedPageIndex.HasValue)
                bookmarkInfo["pageIndex"] = extractedPageIndex.Value;

            bookmarkList.Add(bookmarkInfo);
        }

        var result = new
        {
            count = bookmarkList.Count,
            items = bookmarkList
        };
        return JsonSerializer.Serialize(result, JsonOptions);
    }

    /// <summary>
    ///     Extracts the target page index from a bookmark.
    /// </summary>
    /// <param name="bookmark">The bookmark to extract the page index from.</param>
    /// <param name="document">The PDF document.</param>
    /// <returns>The 1-based page index, or null if not available.</returns>
    private static int? ExtractPageIndex(OutlineItemCollection bookmark, Document document)
    {
        Page? page = null;

        if (bookmark.Destination is ExplicitDestination explicitDest)
            page = explicitDest.Page;
        else if (bookmark.Action is GoToAction { Destination: ExplicitDestination actionDest })
            page = actionDest.Page;

        if (page == null)
            return null;

        var pageIndex = document.Pages.IndexOf(page);
        return pageIndex > 0 ? pageIndex : null;
    }
}