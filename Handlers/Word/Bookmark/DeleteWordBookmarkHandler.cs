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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var keepText = parameters.GetOptional("keepText", true);

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for delete operation");

        var doc = context.Document;

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

        MarkModified(context);

        var remainingCount = doc.Range.Bookmarks.Count;

        var result = $"Bookmark '{name}' deleted successfully\n";
        result += $"Bookmark text: {bookmarkText}\n";
        result += $"Keep text: {(keepText ? "Yes" : "No")}\n";
        result += $"Remaining bookmarks in document: {remainingCount}";

        return result;
    }
}
