using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for editing links in PDF documents.
/// </summary>
public class EditPdfLinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing link annotation on a page in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, linkIndex
    ///     Optional: url, targetPage
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var linkIndex = parameters.GetRequired<int>("linkIndex");
        var url = parameters.GetOptional<string?>("url");
        var targetPage = parameters.GetOptional<int?>("targetPage");

        var doc = context.Document;

        if (pageIndex < 1 || pageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var page = doc.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();

        if (linkIndex < 0 || linkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        var link = links[linkIndex];

        if (!string.IsNullOrEmpty(url))
        {
            link.Action = new GoToURIAction(url);
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > doc.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {doc.Pages.Count}");
            link.Action = new GoToAction(doc.Pages[targetPage.Value]);
        }

        MarkModified(context);

        return Success($"Edited link {linkIndex} on page {pageIndex}.");
    }
}
