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
        var deleteParams = ExtractDeleteParameters(parameters);

        var document = context.Document;

        if (deleteParams.BookmarkIndex < 1 || deleteParams.BookmarkIndex > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[deleteParams.BookmarkIndex];
        var title = bookmark.Title;

        var countBefore = document.Outlines.Count;
        document.Outlines.Delete(title);
        var countAfter = document.Outlines.Count;
        var deletedCount = countBefore - countAfter;

        MarkModified(context);

        return deletedCount > 1
            ? Success(
                $"Deleted {deletedCount} bookmark(s) with title '{title}' (requested index {deleteParams.BookmarkIndex}).")
            : Success($"Deleted bookmark '{title}' (index {deleteParams.BookmarkIndex}).");
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("bookmarkIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete bookmark parameters.
    /// </summary>
    private sealed record DeleteParameters(int BookmarkIndex);
}
