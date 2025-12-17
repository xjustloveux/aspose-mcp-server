using System.Text;
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

        return operation.ToLower() switch
        {
            "add" => await AddBookmarkAsync(arguments, path),
            "edit" => await EditBookmarkAsync(arguments, path),
            "delete" => await DeleteBookmarkAsync(arguments, path),
            "get" => await GetBookmarksAsync(arguments, path),
            "goto" => await GotoBookmarkAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a bookmark to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing name, optional text, paragraphIndex, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var name = ArgumentHelper.GetString(arguments, "name");
        var text = ArgumentHelper.GetStringNullable(arguments, "text");
        var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        // Check if bookmark already exists
        if (doc.Range.Bookmarks[name] != null) throw new InvalidOperationException($"Bookmark '{name}' already exists");

        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning
                if (paragraphs.Count > 0)
                    if (paragraphs[0] is Paragraph firstPara)
                        builder.MoveTo(firstPara);
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                // Insert after the specified paragraph
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

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Edits a bookmark (renames or changes text)
    /// </summary>
    /// <param name="arguments">JSON arguments containing name, optional newName, text or newText, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var bookmarkName = ArgumentHelper.GetString(arguments, "name");
        var newName = ArgumentHelper.GetStringNullable(arguments, "newName");
        // Accept both text and newText for compatibility, text is required for edit operation
        var newText = ArgumentHelper.GetString(arguments, "newText", "text", "newText or text");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        var bookmarks = doc.Range.Bookmarks;

        Bookmark? bookmark;
        try
        {
            bookmark = bookmarks[bookmarkName];
        }
        catch
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
        if (!string.IsNullOrEmpty(newName) && newName != bookmarkName)
        {
            // Check if new name already exists
            Bookmark? existingBookmark = null;
            try
            {
                existingBookmark = bookmarks[newName];
            }
            catch
            {
                // New name doesn't exist, continue
            }

            if (existingBookmark != null) throw new ArgumentException($"Bookmark name '{newName}' already exists");

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
        if (!string.IsNullOrEmpty(newText))
            try
            {
                // Get the bookmark range and replace its content
                var bookmarkRange = bookmark.BookmarkStart;
                var bookmarkEnd = bookmark.BookmarkEnd;

                if (bookmarkRange != null && bookmarkEnd != null)
                {
                    // Remove existing content between bookmark start and end
                    var currentNode = bookmarkRange.NextSibling;
                    while (currentNode != null && currentNode != bookmarkEnd)
                    {
                        var nextNode = currentNode.NextSibling;
                        currentNode.Remove();
                        currentNode = nextNode;
                    }

                    // Insert new text
                    var builder = new DocumentBuilder(doc);
                    builder.MoveTo(bookmarkRange);
                    builder.Write(newText);

                    changes.Add("Bookmark content updated");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to update bookmark content: {ex.Message}", ex);
            }

        if (changes.Count == 0)
            return await Task.FromResult("No changes made. Please provide newName or newText parameter.");

        doc.Save(outputPath);

        var result = $"Bookmark '{bookmarkName}' edited successfully\n";
        result += $"Original name: {oldName}\n";
        result += $"Original content: {oldText}\n";
        if (!string.IsNullOrEmpty(newName)) result += $"New name: {newName}\n";
        if (!string.IsNullOrEmpty(newText)) result += $"New content: {newText}\n";
        result += $"Changes: {string.Join(", ", changes)}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Deletes a bookmark from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing name, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteBookmarkAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var name = ArgumentHelper.GetString(arguments, "name");
        var keepText = ArgumentHelper.GetBool(arguments, "keepText");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);

        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
            throw new ArgumentException($"Bookmark '{name}' not found. Use get operation to view available bookmarks");

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

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Gets all bookmarks from the document
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all bookmarks</returns>
    private async Task<string> GetBookmarksAsync(JsonObject? _, string path)
    {
        var doc = new Document(path);

        // Get all bookmarks
        var bookmarks = doc.Range.Bookmarks;

        if (bookmarks.Count == 0) return await Task.FromResult("No bookmarks found in document");

        var result = new StringBuilder();
        result.AppendLine($"Found {bookmarks.Count} bookmarks:\n");

        var index = 0;
        foreach (var bookmark in bookmarks)
        {
            result.AppendLine($"Bookmark #{index}:");
            result.AppendLine($"  Name: {bookmark.Name}");
            result.AppendLine($"  Text: {bookmark.Text}");
            result.AppendLine($"  Length: {bookmark.Text.Length} characters");
            result.AppendLine();
            index++;
        }

        return await Task.FromResult(result.ToString().TrimEnd());
    }

    /// <summary>
    ///     Gets bookmark content and position
    /// </summary>
    /// <param name="arguments">JSON arguments containing name</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with bookmark information</returns>
    private async Task<string> GotoBookmarkAsync(JsonObject? arguments, string path)
    {
        var name = ArgumentHelper.GetString(arguments, "name");

        var doc = new Document(path);

        // Get bookmark
        var bookmark = doc.Range.Bookmarks[name];
        if (bookmark == null)
            throw new ArgumentException($"Bookmark '{name}' not found. Use get operation to view available bookmarks");

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

        return await Task.FromResult(result);
    }
}