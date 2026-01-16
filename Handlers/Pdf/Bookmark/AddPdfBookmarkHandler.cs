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
        var addParams = ExtractAddParameters(parameters);

        var document = context.Document;

        if (addParams.PageIndex < 1 || addParams.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var bookmark = new OutlineItemCollection(document.Outlines)
        {
            Title = addParams.Title,
            Action = new GoToAction(document.Pages[addParams.PageIndex])
        };

        document.Outlines.Add(bookmark);
        MarkModified(context);

        return Success($"Added bookmark '{addParams.Title}' pointing to page {addParams.PageIndex}.");
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<string>("title"),
            parameters.GetRequired<int>("pageIndex")
        );
    }

    /// <summary>
    ///     Record to hold add bookmark parameters.
    /// </summary>
    private sealed record AddParameters(string Title, int PageIndex);
}
