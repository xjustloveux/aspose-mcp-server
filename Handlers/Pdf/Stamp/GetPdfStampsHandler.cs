using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Stamp;

namespace AsposeMcpServer.Handlers.Pdf.Stamp;

/// <summary>
///     Handler for listing stamp annotations from PDF documents.
/// </summary>
[ResultType(typeof(GetStampsPdfResult))]
public class GetPdfStampsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>
    ///     Lists stamp annotations from one or all pages in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex (default: 0, 0 = all pages).
    /// </param>
    /// <returns>Result containing stamp annotation information.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        var document = context.Document;

        var stamps = new List<PdfStampInfo>();
        var pages = p.PageIndex == 0
            ? Enumerable.Range(1, document.Pages.Count)
            : [p.PageIndex];

        foreach (var pIdx in pages)
        {
            if (pIdx < 1 || pIdx > document.Pages.Count) continue;
            var page = document.Pages[pIdx];
            var annotIdx = 0;
            foreach (var annotation in page.Annotations)
            {
                annotIdx++;
                if (annotation is StampAnnotation stampAnnotation)
                    stamps.Add(new PdfStampInfo
                    {
                        PageIndex = pIdx,
                        Index = annotIdx,
                        Type = "stamp",
                        Contents = stampAnnotation.Contents,
                        Opacity = stampAnnotation.Opacity
                    });
            }
        }

        var pageDesc = p.PageIndex == 0 ? "all pages" : $"page {p.PageIndex}";
        return new GetStampsPdfResult
        {
            Count = stamps.Count,
            Stamps = stamps,
            Message = $"Found {stamps.Count} stamp annotation(s) on {pageDesc}."
        };
    }

    /// <summary>
    ///     Extracts parameters for the list stamps operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetStampsParameters ExtractParameters(OperationParameters parameters)
    {
        return new GetStampsParameters(
            parameters.GetOptional("pageIndex", 0)
        );
    }

    /// <summary>
    ///     Parameters for the list stamps operation.
    /// </summary>
    /// <param name="PageIndex">The page index (0 = all pages, otherwise 1-based).</param>
    private sealed record GetStampsParameters(int PageIndex);
}
