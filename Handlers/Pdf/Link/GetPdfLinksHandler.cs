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
        var pageIndex = parameters.GetOptional<int?>("pageIndex");

        var doc = context.Document;
        List<object> linkList = [];

        if (pageIndex is > 0)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > doc.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

            var page = doc.Pages[pageIndex.Value];
            var links = page.Annotations.OfType<LinkAnnotation>().ToList();

            if (links.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    pageIndex = pageIndex.Value,
                    items = Array.Empty<object>(),
                    message = $"No links found on page {pageIndex.Value}"
                };
                return JsonResult(emptyResult);
            }

            for (var i = 0; i < links.Count; i++)
            {
                var link = links[i];
                var linkInfo = CreateLinkInfo(link, i, pageIndex.Value);
                linkList.Add(linkInfo);
            }

            var result = new
            {
                count = linkList.Count,
                pageIndex = pageIndex.Value,
                items = linkList
            };
            return JsonResult(result);
        }
        else
        {
            for (var p = 1; p <= doc.Pages.Count; p++)
            {
                var page = doc.Pages[p];
                var links = page.Annotations.OfType<LinkAnnotation>().ToList();

                for (var i = 0; i < links.Count; i++)
                {
                    var link = links[i];
                    var linkInfo = CreateLinkInfo(link, i, p);
                    linkList.Add(linkInfo);
                }
            }

            if (linkList.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    items = Array.Empty<object>(),
                    message = "No links found in document"
                };
                return JsonResult(emptyResult);
            }

            var result = new
            {
                count = linkList.Count,
                items = linkList
            };
            return JsonResult(result);
        }
    }

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
}
