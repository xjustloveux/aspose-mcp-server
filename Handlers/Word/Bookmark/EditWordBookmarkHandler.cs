using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for editing bookmarks in Word documents.
/// </summary>
public class EditWordBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing bookmark's name or text content.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: name
    ///     Optional: newName, newText
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var newName = parameters.GetOptional<string?>("newName");
        var newText = parameters.GetOptional<string?>("newText");

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for edit operation");
        if (string.IsNullOrEmpty(newText) && string.IsNullOrEmpty(newName))
            throw new ArgumentException("newName or newText is required for edit operation");

        var doc = context.Document;
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

        MarkModified(context);

        var result = $"Bookmark '{name}' edited successfully\n";
        result += $"Original name: {oldName}\n";
        result += $"Original content: {oldText}\n";
        if (!string.IsNullOrEmpty(newName)) result += $"New name: {newName}\n";
        if (!string.IsNullOrEmpty(newText)) result += $"New content: {newText}\n";
        result += $"Changes: {string.Join(", ", changes)}";

        return result;
    }
}
