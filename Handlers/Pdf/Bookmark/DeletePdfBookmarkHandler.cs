using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Bookmark;

/// <summary>
///     Handler for deleting bookmarks from PDF documents.
/// </summary>
public class DeletePdfBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a bookmark from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: bookmarkIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var bookmarkIndex = parameters.GetRequired<int>("bookmarkIndex");

        var document = context.Document;

        if (bookmarkIndex < 1 || bookmarkIndex > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[bookmarkIndex];
        var title = bookmark.Title;

        var countBefore = document.Outlines.Count;
        document.Outlines.Delete(title);
        var countAfter = document.Outlines.Count;
        var deletedCount = countBefore - countAfter;

        MarkModified(context);

        return deletedCount > 1
            ? Success($"Deleted {deletedCount} bookmark(s) with title '{title}' (requested index {bookmarkIndex}).")
            : Success($"Deleted bookmark '{title}' (index {bookmarkIndex}).");
    }
}
