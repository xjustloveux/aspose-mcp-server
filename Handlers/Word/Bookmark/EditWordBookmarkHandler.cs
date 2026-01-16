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
        var p = ExtractEditParameters(parameters);

        var doc = context.Document;
        var bookmarks = doc.Range.Bookmarks;

        var bookmark = bookmarks[p.Name];
        if (bookmark == null)
        {
            var availableBookmarks = bookmarks.Select(b => b.Name).Take(10).ToList();
            var availableInfo = GetAvailableBookmarksInfo(availableBookmarks, bookmarks.Count);
            throw new ArgumentException(
                $"Bookmark '{p.Name}' not found{availableInfo}. Use get operation to view all available bookmarks");
        }

        var oldName = bookmark.Name;
        var oldText = bookmark.Text;
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.NewName) && !p.NewName.Equals(p.Name, StringComparison.OrdinalIgnoreCase))
        {
            var existingWithNewName = bookmarks
                .FirstOrDefault(b => b.Name.Equals(p.NewName, StringComparison.OrdinalIgnoreCase));
            if (existingWithNewName != null)
                throw new ArgumentException(
                    $"Bookmark name '{existingWithNewName.Name}' already exists (bookmark names are case-insensitive)");

            bookmark.Name = p.NewName;
            changes.Add($"Bookmark name: {oldName} -> {p.NewName}");
        }

        if (!string.IsNullOrEmpty(p.NewText))
        {
            bookmark.Text = p.NewText;
            changes.Add("Bookmark content updated");
        }

        if (changes.Count == 0)
            return "No changes made. Please provide newName or newText parameter.";

        MarkModified(context);

        var result = $"Bookmark '{p.Name}' edited successfully\n";
        result += $"Original name: {oldName}\n";
        result += $"Original content: {oldText}\n";
        if (!string.IsNullOrEmpty(p.NewName)) result += $"New name: {p.NewName}\n";
        if (!string.IsNullOrEmpty(p.NewText)) result += $"New content: {p.NewText}\n";
        result += $"Changes: {string.Join(", ", changes)}";

        return result;
    }

    /// <summary>
    ///     Gets the available bookmarks info string.
    /// </summary>
    /// <param name="availableBookmarks">The list of available bookmark names.</param>
    /// <param name="totalCount">The total count of bookmarks.</param>
    /// <returns>The info string about available bookmarks.</returns>
    private static string GetAvailableBookmarksInfo(List<string> availableBookmarks, int totalCount)
    {
        if (availableBookmarks.Count == 0)
            return "\nDocument has no bookmarks";

        var suffix = totalCount > 10 ? "..." : "";
        return $"\nAvailable bookmarks: {string.Join(", ", availableBookmarks)}{suffix}";
    }

    /// <summary>
    ///     Extracts and validates parameters for the edit bookmark operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are not provided.</exception>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var newName = parameters.GetOptional<string?>("newName");
        var newText = parameters.GetOptional<string?>("newText");

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for edit operation");
        if (string.IsNullOrEmpty(newText) && string.IsNullOrEmpty(newName))
            throw new ArgumentException("newName or newText is required for edit operation");

        return new EditParameters(name, newName, newText);
    }

    /// <summary>
    ///     Parameters for the edit bookmark operation.
    /// </summary>
    /// <param name="Name">The bookmark name to edit.</param>
    /// <param name="NewName">The new name for the bookmark.</param>
    /// <param name="NewText">The new text content for the bookmark.</param>
    private sealed record EditParameters(string Name, string? NewName, string? NewText);
}
