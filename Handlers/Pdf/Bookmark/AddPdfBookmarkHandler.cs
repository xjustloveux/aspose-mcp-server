using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Bookmark;

/// <summary>
///     Handler for adding bookmarks to PDF documents.
/// </summary>
public class AddPdfBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new bookmark to the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: title, pageIndex
    /// </param>
    /// <returns>Success message with bookmark details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var title = parameters.GetRequired<string>("title");
        var pageIndex = parameters.GetRequired<int>("pageIndex");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var bookmark = new OutlineItemCollection(document.Outlines)
        {
            Title = title,
            Action = new GoToAction(document.Pages[pageIndex])
        };

        document.Outlines.Add(bookmark);
        MarkModified(context);

        return Success($"Added bookmark '{title}' pointing to page {pageIndex}.");
    }
}
