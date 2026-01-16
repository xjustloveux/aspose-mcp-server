using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for getting links from PDF documents.
/// </summary>
public class GetPdfLinksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all link annotations from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex (if specified, gets links from that page only)
    /// </param>
    /// <returns>JSON string containing link information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);
        var doc = context.Document;

        return p.PageIndex is > 0
            ? GetLinksFromPage(doc, p.PageIndex.Value)
            : GetLinksFromDocument(doc);
    }

    /// <summary>
    ///     Extracts get parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(
            parameters.GetOptional<int?>("pageIndex"));
    }

    /// <summary>
    ///     Retrieves link annotations from a specific page.
    /// </summary>
    /// <param name="doc">The PDF document.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <returns>JSON string containing link information from the specified page.</returns>
    private string GetLinksFromPage(Document doc, int pageIndex)
    {
        if (pageIndex < 1 || pageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var links = doc.Pages[pageIndex].Annotations.OfType<LinkAnnotation>().ToList();

        if (links.Count == 0)
            return JsonResult(new
            {
                count = 0,
                pageIndex,
                items = Array.Empty<object>(),
                message = $"No links found on page {pageIndex}"
            });

        var linkList = links.Select((link, i) => CreateLinkInfo(link, i, pageIndex)).ToList();

        return JsonResult(new { count = linkList.Count, pageIndex, items = linkList });
    }

    /// <summary>
    ///     Retrieves link annotations from all pages in the document.
    /// </summary>
    /// <param name="doc">The PDF document.</param>
    /// <returns>JSON string containing link information from all pages.</returns>
    private string GetLinksFromDocument(Document doc)
    {
        List<object> linkList = [];

        for (var p = 1; p <= doc.Pages.Count; p++)
        {
            var links = doc.Pages[p].Annotations.OfType<LinkAnnotation>().ToList();
            linkList.AddRange(links.Select((link, i) => CreateLinkInfo(link, i, p)));
        }

        if (linkList.Count == 0)
            return JsonResult(new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No links found in document"
            });

        return JsonResult(new { count = linkList.Count, items = linkList });
    }

    /// <summary>
    ///     Creates a dictionary containing link annotation information.
    /// </summary>
    /// <param name="link">The link annotation.</param>
    /// <param name="index">The 0-based link index within the page.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <returns>A dictionary containing link information.</returns>
    private static Dictionary<string, object?> CreateLinkInfo(LinkAnnotation link, int index, int pageIndex)
    {
        var linkInfo = new Dictionary<string, object?>
        {
            ["index"] = index,
            ["pageIndex"] = pageIndex,
            ["x"] = link.Rect.LLX,
            ["y"] = link.Rect.LLY
        };

        if (link.Action is GoToURIAction uriAction)
        {
            linkInfo["type"] = "url";
            linkInfo["url"] = uriAction.URI;
        }
        else if (link.Action is GoToAction gotoAction)
        {
            linkInfo["type"] = "page";
            if (gotoAction.Destination is XYZExplicitDestination xyzDest)
                linkInfo["destinationPage"] = xyzDest.PageNumber;
            else if (gotoAction.Destination is ExplicitDestination explicitDest)
                linkInfo["destinationPage"] = explicitDest.PageNumber;
        }

        return linkInfo;
    }

    /// <summary>
    ///     Parameters for getting links.
    /// </summary>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    private sealed record GetParameters(int? PageIndex);
}
