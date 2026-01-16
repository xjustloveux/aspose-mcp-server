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
        var p = ExtractEditParameters(parameters);

        var doc = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var page = doc.Pages[p.PageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();

        if (p.LinkIndex < 0 || p.LinkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        var link = links[p.LinkIndex];

        if (!string.IsNullOrEmpty(p.Url))
        {
            link.Action = new GoToURIAction(p.Url);
        }
        else if (p.TargetPage.HasValue)
        {
            if (p.TargetPage.Value < 1 || p.TargetPage.Value > doc.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {doc.Pages.Count}");
            link.Action = new GoToAction(doc.Pages[p.TargetPage.Value]);
        }

        MarkModified(context);

        return Success($"Edited link {p.LinkIndex} on page {p.PageIndex}.");
    }

    /// <summary>
    ///     Extracts edit parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<int>("linkIndex"),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<int?>("targetPage"));
    }

    /// <summary>
    ///     Parameters for editing a link.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="LinkIndex">The 0-based link index.</param>
    /// <param name="Url">The optional URL for external links.</param>
    /// <param name="TargetPage">The optional target page for internal links.</param>
    private record EditParameters(int PageIndex, int LinkIndex, string? Url, int? TargetPage);
}
