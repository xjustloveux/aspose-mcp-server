using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Bookmark;

/// <summary>
///     Handler for editing bookmarks in PDF documents.
/// </summary>
public class EditPdfBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing bookmark in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: bookmarkIndex
    ///     Optional: title, pageIndex
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var bookmarkIndex = parameters.GetRequired<int>("bookmarkIndex");
        var title = parameters.GetOptional<string?>("title");
        var pageIndex = parameters.GetOptional<int?>("pageIndex");

        var document = context.Document;

        if (bookmarkIndex < 1 || bookmarkIndex > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[bookmarkIndex];

        if (!string.IsNullOrEmpty(title))
            bookmark.Title = title;

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            bookmark.Action = new GoToAction(document.Pages[pageIndex.Value]);
        }

        MarkModified(context);

        return Success($"Edited bookmark (index {bookmarkIndex}).");
    }
}
