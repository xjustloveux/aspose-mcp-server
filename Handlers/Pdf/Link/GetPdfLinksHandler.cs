using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Link;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Handler for getting links from PDF documents.
/// </summary>
[ResultType(typeof(GetLinksResult))]
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
    /// <returns>A GetLinksResult containing link information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
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
    /// <returns>GetLinksResult containing link information from the specified page.</returns>
    private static GetLinksResult GetLinksFromPage(Document doc, int pageIndex)
    {
        if (pageIndex < 1 || pageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var links = doc.Pages[pageIndex].Annotations.OfType<LinkAnnotation>().ToList();

        if (links.Count == 0)
            return new GetLinksResult
            {
                Count = 0,
                PageIndex = pageIndex,
                Items = [],
                Message = $"No links found on page {pageIndex}"
            };

        var linkList = links.Select((link, i) => CreateLinkInfo(link, i, pageIndex)).ToList();

        return new GetLinksResult
        {
            Count = linkList.Count,
            PageIndex = pageIndex,
            Items = linkList
        };
    }

    /// <summary>
    ///     Retrieves link annotations from all pages in the document.
    /// </summary>
    /// <param name="doc">The PDF document.</param>
    /// <returns>GetLinksResult containing link information from all pages.</returns>
    private static GetLinksResult GetLinksFromDocument(Document doc)
    {
        List<LinkInfo> linkList = [];

        for (var p = 1; p <= doc.Pages.Count; p++)
        {
            var links = doc.Pages[p].Annotations.OfType<LinkAnnotation>().ToList();
            linkList.AddRange(links.Select((link, i) => CreateLinkInfo(link, i, p)));
        }

        if (linkList.Count == 0)
            return new GetLinksResult
            {
                Count = 0,
                Items = [],
                Message = "No links found in document"
            };

        return new GetLinksResult
        {
            Count = linkList.Count,
            Items = linkList
        };
    }

    /// <summary>
    ///     Creates a LinkInfo object containing link annotation information.
    /// </summary>
    /// <param name="link">The link annotation.</param>
    /// <param name="index">The 0-based link index within the page.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <returns>A LinkInfo object containing link information.</returns>
    private static LinkInfo CreateLinkInfo(LinkAnnotation link, int index, int pageIndex)
    {
        string? type = null;
        string? url = null;
        int? destinationPage = null;

        if (link.Action is GoToURIAction uriAction)
        {
            type = "url";
            url = uriAction.URI;
        }
        else if (link.Action is GoToAction gotoAction)
        {
            type = "page";
            if (gotoAction.Destination is XYZExplicitDestination xyzDest)
                destinationPage = xyzDest.PageNumber;
            else if (gotoAction.Destination is ExplicitDestination explicitDest)
                destinationPage = explicitDest.PageNumber;
        }

        return new LinkInfo
        {
            Index = index,
            PageIndex = pageIndex,
            X = link.Rect.LLX,
            Y = link.Rect.LLY,
            Type = type,
            Url = url,
            DestinationPage = destinationPage
        };
    }

    /// <summary>
    ///     Parameters for getting links.
    /// </summary>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    private sealed record GetParameters(int? PageIndex);
}
