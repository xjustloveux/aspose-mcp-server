using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word bookmarks (add, edit, delete, get, goto)
///     Merges: WordAddBookmarkTool, WordEditBookmarkTool, WordDeleteBookmarkTool, WordGetBookmarksTool,
///     WordGotoBookmarkTool
/// </summary>
public class WordBookmarkTool : IAsposeTool
{
    public string Description => @"Manage Word bookmarks. Supports 5 operations: add, edit, delete, get, goto.

Usage examples:
- Add bookmark: word_bookmark(operation='add', path='doc.docx', name='bookmark1', text='Bookmarked text')
- Edit bookmark: word_bookmark(operation='edit', path='doc.docx', name='bookmark1', text='Updated text')
- Delete bookmark: word_bookmark(operation='delete', path='doc.docx', name='bookmark1')
- Get bookmarks: word_bookmark(operation='get', path='doc.docx')
- Goto bookmark: word_bookmark(operation='goto', path='doc.docx', name='bookmark1')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a bookmark (required params: path, name)
- 'edit': Edit a bookmark (required params: path, name, newText or text)
- 'delete': Delete a bookmark (required params: path, name)
- 'get': Get all bookmarks (required params: path)
- 'goto': Get bookmark location (required params: path, name)",
                @enum = new[] { "add", "edit", "delete", "get", "goto" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/edit/delete operations)"
            },
            name = new
            {
                type = "string",
                description = "Bookmark name (required for add/edit/delete/goto operations)"
            },
            text = new
            {
                type = "string",
                description = "Text content to bookmark (optional, for add operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert bookmark at (0-based, optional, for add operation)"
            },
            newName = new
            {
                type = "string",
                description = "New bookmark name (optional, for edit operation)"
            },
            newText = new
            {
                type = "string",
                description =
                    "New text content for the bookmark (required for edit operation, can also use 'text' parameter)"
            },
            keepText = new
            {
                type = "boolean",
                description = "Keep bookmark text content when deleting (default: true, for delete operation)"
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
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add" => await AddBookmarkAsync(path, outputPath, arguments),
            "edit" => await EditBookmarkAsync(path, outputPath, arguments),
            "delete" => await DeleteBookmarkAsync(path, outputPath, arguments),
            "get" => await GetBookmarksAsync(path),
            "goto" => await GotoBookmarkAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a bookmark to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing name, optional text, paragraphIndex</param>
    /// <returns>Success message</returns>
    private Task<string> AddBookmarkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");
            var text = ArgumentHelper.GetStringNullable(arguments, "text");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");

            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            // Check if bookmark already exists (case-insensitive in Word)
            var existingBookmark = doc.Range.Bookmarks
                .FirstOrDefault(b => b.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (existingBookmark != null)
                throw new InvalidOperationException(
                    $"Bookmark '{existingBookmark.Name}' already exists (bookmark names are case-insensitive)");

            // Determine insertion position
            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

                // Protection for empty document
                if (paragraphs.Count == 0)
                {
                    // Document is empty, move to document end (will create content there)
                    builder.MoveToDocumentEnd();
                }
                else if (paragraphIndex.Value == -1)
                {
                    // Insert at the beginning
                    if (paragraphs[0] is Paragraph firstPara)
                        builder.MoveTo(firstPara);
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    // Insert after the specified paragraph
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                        builder.MoveTo(targetPara);
                    else
                        throw new InvalidOperationException(
                            $"Unable to find paragraph at index {paragraphIndex.Value}");
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
                }
            }
            else
            {
                // Default: Move to end of document
                builder.MoveToDocumentEnd();
            }

            // Insert bookmark
            builder.StartBookmark(name);

            // Add text if provided
            if (!string.IsNullOrEmpty(text)) builder.Write(text);

            builder.EndBookmark(name);

            doc.Save(outputPath);

            var result = "Bookmark added successfully\n";
            result += $"Bookmark name: {name}\n";
            if (!string.IsNullOrEmpty(text)) result += $"Bookmark text: {text}\n";
            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                    result += "Insert position: beginning of document\n";
                else
                    result += $"Insert position: after paragraph #{paragraphIndex.Value}\n";
            }
            else
            {
                result += "Insert position: end of document\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits a bookmark (renames or changes text)
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing name, optional newName, newText or text</param>
    /// <returns>Success message</returns>
    private Task<string> EditBookmarkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var bookmarkName = ArgumentHelper.GetString(arguments, "name");
            var newName = ArgumentHelper.GetStringNullable(arguments, "newName");
            // Accept both text and newText for compatibility, text is required for edit operation
            var newText = ArgumentHelper.GetString(arguments, "newText", "text", "newText or text");

            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var doc = new Document(path);
            var bookmarks = doc.Range.Bookmarks;

            Bookmark? bookmark;
            try
            {
                bookmark = bookmarks[bookmarkName];
            }
            catch (Exception ex)
            {
                // Get available bookmarks for better error message
                Console.Error.WriteLine($"[WARN] Error accessing bookmark '{bookmarkName}': {ex.Message}");
                var availableBookmarks = new List<string>();
                foreach (var bm in bookmarks) availableBookmarks.Add(bm.Name);
                var availableInfo = availableBookmarks.Count > 0
                    ? $"\nAvailable bookmarks: {string.Join(", ", availableBookmarks.Take(10))}{(availableBookmarks.Count > 10 ? "..." : "")}"
                    : "\nDocument has no bookmarks";
                throw new ArgumentException(
                    $"Bookmark '{bookmarkName}' not found{availableInfo}. Use get operation to view all available bookmarks");
            }

            if (bookmark == null)
            {
                // Get available bookmarks for better error message
                var availableBookmarks = new List<string>();
                foreach (var bm in bookmarks) availableBookmarks.Add(bm.Name);
                var availableInfo = availableBookmarks.Count > 0
                    ? $"\nAvailable bookmarks: {string.Join(", ", availableBookmarks.Take(10))}{(availableBookmarks.Count > 10 ? "..." : "")}"
                    : "\nDocument has no bookmarks";
                throw new ArgumentException(
                    $"Bookmark '{bookmarkName}' not found{availableInfo}. Use get operation to view all available bookmarks");
            }

            var oldName = bookmark.Name;
            var oldText = bookmark.Text;
            var changes = new List<string>();

            // Rename bookmark if new name provided
            if (!string.IsNullOrEmpty(newName) && !newName.Equals(bookmarkName, StringComparison.OrdinalIgnoreCase))
            {
                // Check if new name already exists (case-insensitive in Word)
                var existingWithNewName = bookmarks
                    .FirstOrDefault(b => b.Name.Equals(newName, StringComparison.OrdinalIgnoreCase));
                if (existingWithNewName != null)
                    throw new ArgumentException(
                        $"Bookmark name '{existingWithNewName.Name}' already exists (bookmark names are case-insensitive)");

                try
                {
                    bookmark.Name = newName;
                    changes.Add($"Bookmark name: {oldName} -> {newName}");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Unable to rename bookmark: {ex.Message}", ex);
                }
            }

            // Update bookmark text if new text provided
            // Use Aspose's built-in Text property which handles node deletion and insertion automatically
            if (!string.IsNullOrEmpty(newText))
            {
                bookmark.Text = newText;
                changes.Add("Bookmark content updated");
            }

            if (changes.Count == 0)
                return "No changes made. Please provide newName or newText parameter.";

            doc.Save(outputPath);

            var result = $"Bookmark '{bookmarkName}' edited successfully\n";
            result += $"Original name: {oldName}\n";
            result += $"Original content: {oldText}\n";
            if (!string.IsNullOrEmpty(newName)) result += $"New name: {newName}\n";
            if (!string.IsNullOrEmpty(newText)) result += $"New content: {newText}\n";
            result += $"Changes: {string.Join(", ", changes)}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a bookmark from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing name, optional keepText</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteBookmarkAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");
            var keepText = ArgumentHelper.GetBool(arguments, "keepText");

            var outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            var doc = new Document(path);

            // Get bookmark
            var bookmark = doc.Range.Bookmarks[name];
            if (bookmark == null)
                throw new ArgumentException(
                    $"Bookmark '{name}' not found. Use get operation to view available bookmarks");

            // Get bookmark info before deletion
            var bookmarkText = bookmark.Text;

            // Delete bookmark
            if (keepText)
            {
                // Remove bookmark markers but keep text
                bookmark.Remove();
            }
            else
            {
                // Remove bookmark and its text content
                bookmark.BookmarkStart?.Remove();
                bookmark.BookmarkEnd?.Remove();
            }

            doc.Save(outputPath);

            // Count remaining bookmarks
            var remainingCount = doc.Range.Bookmarks.Count;

            var result = $"Bookmark '{name}' deleted successfully\n";
            result += $"Bookmark text: {bookmarkText}\n";
            result += $"Keep text: {(keepText ? "Yes" : "No")}\n";
            result += $"Remaining bookmarks in document: {remainingCount}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all bookmarks from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <returns>JSON formatted string with all bookmarks for better LLM processing</returns>
    private Task<string> GetBookmarksAsync(string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);

            // Get all bookmarks
            var bookmarks = doc.Range.Bookmarks;

            if (bookmarks.Count == 0)
                return JsonSerializer.Serialize(new
                    { count = 0, bookmarks = Array.Empty<object>(), message = "No bookmarks found in document" });

            // Build JSON response for better LLM processing
            var bookmarkList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Gets bookmark content and position
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing name</param>
    /// <returns>Formatted string with bookmark information</returns>
    private Task<string> GotoBookmarkAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");

            var doc = new Document(path);

            // Get bookmark
            var bookmark = doc.Range.Bookmarks[name];
            if (bookmark == null)
                throw new ArgumentException(
                    $"Bookmark '{name}' not found. Use get operation to view available bookmarks");

            // Get bookmark information
            var bookmarkText = bookmark.Text;
            var bookmarkRange = bookmark.BookmarkStart?.ParentNode as Paragraph;

            // Try to find paragraph index
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

            // Get surrounding context if possible
            if (bookmarkRange != null)
            {
                var paraText = bookmarkRange.GetText().Trim();
                if (paraText.Length > 100) paraText = paraText.Substring(0, 100) + "...";
                result += $"Paragraph content: {paraText}\n";
            }

            return result;
        });
    }
}