using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for deleting bookmarks from Word documents.
/// </summary>
public class DeleteWordBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a bookmark from the document, optionally keeping its text content.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: name
    ///     Optional: keepText (default: true)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided or bookmark is not found.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteParameters(parameters);

        var doc = context.Document;

        var bookmark = doc.Range.Bookmarks[p.Name];
        if (bookmark == null)
            throw new ArgumentException(
                $"Bookmark '{p.Name}' not found. Use get operation to view available bookmarks");

        var bookmarkText = bookmark.Text;

        if (p.KeepText)
        {
            bookmark.Remove();
        }
        else
        {
            bookmark.BookmarkStart?.Remove();
            bookmark.BookmarkEnd?.Remove();
        }

        MarkModified(context);

        var remainingCount = doc.Range.Bookmarks.Count;

        var result = $"Bookmark '{p.Name}' deleted successfully\n";
        result += $"Bookmark text: {bookmarkText}\n";
        result += $"Keep text: {(p.KeepText ? "Yes" : "No")}\n";
        result += $"Remaining bookmarks in document: {remainingCount}";

        return result;
    }

    /// <summary>
    ///     Extracts and validates parameters for the delete bookmark operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided.</exception>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var keepText = parameters.GetOptional("keepText", true);

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for delete operation");

        return new DeleteParameters(name, keepText);
    }

    /// <summary>
    ///     Parameters for the delete bookmark operation.
    /// </summary>
    /// <param name="Name">The bookmark name to delete.</param>
    /// <param name="KeepText">Whether to keep the bookmark text content.</param>
    private sealed record DeleteParameters(string Name, bool KeepText);
}
