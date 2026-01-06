using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word bookmarks (add, edit, delete, get, goto)
/// </summary>
[McpServerToolType]
public class WordBookmarkTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordBookmarkTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordBookmarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word bookmark operation (add, edit, delete, get, goto).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, goto.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="name">Bookmark name.</param>
    /// <param name="text">Text content for bookmark.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for beginning).</param>
    /// <param name="newName">New bookmark name (for edit).</param>
    /// <param name="newText">New text content (for edit).</param>
    /// <param name="keepText">Keep text when deleting (default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_bookmark")]
    [Description(@"Manage Word bookmarks. Supports 5 operations: add, edit, delete, get, goto.

Usage examples:
- Add bookmark: word_bookmark(operation='add', path='doc.docx', name='bookmark1', text='Bookmarked text')
- Edit bookmark: word_bookmark(operation='edit', path='doc.docx', name='bookmark1', text='Updated text')
- Delete bookmark: word_bookmark(operation='delete', path='doc.docx', name='bookmark1')
- Get bookmarks: word_bookmark(operation='get', path='doc.docx')
- Goto bookmark: word_bookmark(operation='goto', path='doc.docx', name='bookmark1')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get, goto")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Bookmark name")] string? name = null,
        [Description("Text content for bookmark")]
        string? text = null,
        [Description("Paragraph index (0-based, -1 for beginning)")]
        int? paragraphIndex = null,
        [Description("New bookmark name (for edit)")]
        string? newName = null,
        [Description("New text content (for edit)")]
        string? newText = null,
        [Description("Keep text when deleting (default: true)")]
        bool keepText = true)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddBookmark(ctx, outputPath, name, text, paragraphIndex),
            "edit" => EditBookmark(ctx, outputPath, name, newName, newText ?? text),
            "delete" => DeleteBookmark(ctx, outputPath, name, keepText),
            "get" => GetBookmarks(ctx),
            "goto" => GotoBookmark(ctx, name),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a bookmark to the document at the specified location.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="name">The bookmark name.</param>
    /// <param name="text">The text content for the bookmark.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based, -1 for beginning).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the bookmark name is empty or paragraph index is invalid.</exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when a bookmark with the same name already exists or paragraph
    ///     cannot be found.
    /// </exception>
    private static string AddBookmark(DocumentContext<Document> ctx, string? outputPath, string? name, string? text,
        int? paragraphIndex)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for add operation");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        // Check if bookmark already exists
        var existingBookmark = doc.Range.Bookmarks
            .FirstOrDefault(b => b.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existingBookmark != null)
            throw new InvalidOperationException(
                $"Bookmark '{existingBookmark.Name}' already exists (bookmark names are case-insensitive)");

        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphs.Count == 0)
            {
                builder.MoveToDocumentEnd();
            }
            else if (paragraphIndex.Value == -1)
            {
                if (paragraphs[0] is Paragraph firstPara)
                    builder.MoveTo(firstPara);
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                    builder.MoveTo(targetPara);
                else
                    throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.StartBookmark(name);
        if (!string.IsNullOrEmpty(text)) builder.Write(text);
        builder.EndBookmark(name);

        ctx.Save(outputPath);

        var result = "Bookmark added successfully\n";
        result += $"Bookmark name: {name}\n";
        if (!string.IsNullOrEmpty(text)) result += $"Bookmark text: {text}\n";
        if (paragraphIndex.HasValue)
            result += paragraphIndex.Value == -1
                ? "Insert position: beginning of document\n"
                : $"Insert position: after paragraph #{paragraphIndex.Value}\n";
        else
            result += "Insert position: end of document\n";

        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits an existing bookmark's name or text content.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="name">The current bookmark name.</param>
    /// <param name="newName">The new bookmark name (optional).</param>
    /// <param name="newText">The new text content (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the bookmark name is empty, bookmark not found, or new name already
    ///     exists.
    /// </exception>
    private static string EditBookmark(DocumentContext<Document> ctx, string? outputPath, string? name, string? newName,
        string? newText)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for edit operation");
        if (string.IsNullOrEmpty(newText) && string.IsNullOrEmpty(newName))
            throw new ArgumentException("newName or newText is required for edit operation");

        var doc = ctx.Document;
        var bookmarks = doc.Range.Bookmarks;

        var bookmark = bookmarks[name];
        if (bookmark == null)
        {
            var availableBookmarks = bookmarks.Select(b => b.Name).Take(10).ToList();
            var availableInfo = availableBookmarks.Count > 0
                ? $"\nAvailable bookmarks: {string.Join(", ", availableBookmarks)}{(bookmarks.Count > 10 ? "..." : "")}"
                : "\nDocument has no bookmarks";
            throw new ArgumentException(
                $"Bookmark '{name}' not found{availableInfo}. Use get operation to view all available bookmarks");
        }

        var oldName = bookmark.Name;
        var oldText = bookmark.Text;
        List<string> changes = [];

        if (!string.IsNullOrEmpty(newName) && !newName.Equals(name, StringComparison.OrdinalIgnoreCase))
        {
            var existingWithNewName = bookmarks
                .FirstOrDefault(b => b.Name.Equals(newName, StringComparison.OrdinalIgnoreCase));
            if (existingWithNewName != null)
                throw new ArgumentException(
                    $"Bookmark name '{existingWithNewName.Name}' already exists (bookmark names are case-insensitive)");

            bookmark.Name = newName;
            changes.Add($"Bookmark name: {oldName} -> {newName}");
        }

        if (!string.IsNullOrEmpty(newText))
        {
            bookmark.Text = newText;
            changes.Add("Bookmark content updated");
        }

        if (changes.Count == 0)
            return "No changes made. Please provide newName or newText parameter.";

        ctx.Save(outputPath);

        var result = $"Bookmark '{name}' edited successfully\n";
        result += $"Original name: {oldName}\n";
        result += $"Original content: {oldText}\n";
        if (!string.IsNullOrEmpty(newName)) result += $"New name: {newName}\n";
        if (!string.IsNullOrEmpty(newText)) result += $"New content: {newText}\n";
        result += $"Changes: {string.Join(", ", changes)}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes a bookmark from the document, optionally keeping its text content.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="name">The bookmark name to delete.</param>
    /// <param name="keepText">Whether to keep the bookmark's text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the bookmark name is empty or bookmark not found.</exception>
    private static string DeleteBookmark(DocumentContext<Document> ctx, string? outputPath, string? name, bool keepText)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for delete operation");

        var doc = ctx.Document;

        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
            throw new ArgumentException(
                $"Bookmark '{name}' not found. Use get operation to view available bookmarks");

        var bookmarkText = bookmark.Text;

        if (keepText)
        {
            bookmark.Remove();
        }
        else
        {
            bookmark.BookmarkStart?.Remove();
            bookmark.BookmarkEnd?.Remove();
        }

        ctx.Save(outputPath);

        var remainingCount = doc.Range.Bookmarks.Count;

        var result = $"Bookmark '{name}' deleted successfully\n";
        result += $"Bookmark text: {bookmarkText}\n";
        result += $"Keep text: {(keepText ? "Yes" : "No")}\n";
        result += $"Remaining bookmarks in document: {remainingCount}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Gets all bookmarks from the document as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing all bookmarks information.</returns>
    private static string GetBookmarks(DocumentContext<Document> ctx)
    {
        var doc = ctx.Document;
        var bookmarks = doc.Range.Bookmarks;

        if (bookmarks.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, bookmarks = Array.Empty<object>(), message = "No bookmarks found in document" });

        List<object> bookmarkList = [];
        var index = 0;
        foreach (var bookmark in bookmarks)
        {
            bookmarkList.Add(new
            {
                index,
                name = bookmark.Name,
                text = bookmark.Text,
                length = bookmark.Text.Length
            });
            index++;
        }

        var result = new
        {
            count = bookmarks.Count,
            bookmarks = bookmarkList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets location information for a specific bookmark.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="name">The bookmark name.</param>
    /// <returns>A message containing the bookmark's location information.</returns>
    /// <exception cref="ArgumentException">Thrown when the bookmark name is empty or bookmark not found.</exception>
    private static string GotoBookmark(DocumentContext<Document> ctx, string? name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for goto operation");

        var doc = ctx.Document;

        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
            throw new ArgumentException(
                $"Bookmark '{name}' not found. Use get operation to view available bookmarks");

        var bookmarkText = bookmark.Text;
        var bookmarkRange = bookmark.BookmarkStart?.ParentNode as Paragraph;

        var paragraphIndex = -1;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (var i = 0; i < paragraphs.Count; i++)
            if (paragraphs[i] == bookmarkRange)
            {
                paragraphIndex = i;
                break;
            }

        var result = "Bookmark location information\n";
        result += $"Bookmark name: {name}\n";
        result += $"Bookmark text: {bookmarkText}\n";
        if (paragraphIndex >= 0) result += $"Paragraph index: {paragraphIndex}\n";
        result += $"Bookmark range length: {bookmarkText.Length} characters\n";

        if (bookmarkRange != null)
        {
            var paraText = bookmarkRange.GetText().Trim();
            if (paraText.Length > 100) paraText = paraText[..100] + "...";
            result += $"Paragraph content: {paraText}\n";
        }

        return result;
    }
}