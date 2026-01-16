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
        var editParams = ExtractEditParameters(parameters);

        var document = context.Document;

        if (editParams.BookmarkIndex < 1 || editParams.BookmarkIndex > document.Outlines.Count)
            throw new ArgumentException($"bookmarkIndex must be between 1 and {document.Outlines.Count}");

        var bookmark = document.Outlines[editParams.BookmarkIndex];

        if (!string.IsNullOrEmpty(editParams.Title))
            bookmark.Title = editParams.Title;

        if (editParams.PageIndex.HasValue)
        {
            if (editParams.PageIndex.Value < 1 || editParams.PageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            bookmark.Action = new GoToAction(document.Pages[editParams.PageIndex.Value]);
        }

        MarkModified(context);

        return Success($"Edited bookmark (index {editParams.BookmarkIndex}).");
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<int>("bookmarkIndex"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<int?>("pageIndex")
        );
    }

    /// <summary>
    ///     Record to hold edit bookmark parameters.
    /// </summary>
    private sealed record EditParameters(int BookmarkIndex, string? Title, int? PageIndex);
}
