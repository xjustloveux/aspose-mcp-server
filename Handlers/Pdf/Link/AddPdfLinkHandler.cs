using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for adding links to PDF documents.
/// </summary>
public class AddPdfLinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a link to a specific page in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, url or targetPage.
    ///     Optional: x, y, width, height.
    /// </param>
    /// <returns>Success message with link details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var url = parameters.GetOptional<string?>("url");
        var targetPage = parameters.GetOptional<int?>("targetPage");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 700.0);
        var width = parameters.GetOptional("width", 100.0);
        var height = parameters.GetOptional("height", 20.0);

        if (string.IsNullOrEmpty(url) && !targetPage.HasValue)
            throw new ArgumentException("Either 'url' or 'targetPage' is required");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);

        var linkAnnotation = new LinkAnnotation(page, rect);

        if (!string.IsNullOrEmpty(url))
        {
            linkAnnotation.Action = new GoToURIAction(url);
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");

            linkAnnotation.Action = new GoToAction(document.Pages[targetPage.Value]);
        }

        page.Annotations.Add(linkAnnotation);

        MarkModified(context);

        var linkType = !string.IsNullOrEmpty(url) ? $"URL: {url}" : $"Page: {targetPage}";
        return Success($"Link added to page {pageIndex} ({linkType}).");
    }
}
