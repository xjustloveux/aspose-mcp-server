using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Page;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for getting detailed information about a specific page in PDF documents.
/// </summary>
[ResultType(typeof(GetPdfPageDetailsResult))]
public class GetPdfPageDetailsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_details";

    /// <summary>
    ///     Gets detailed information about a specific page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based page index)
    /// </param>
    /// <returns>JSON string containing detailed page information.</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetDetailsParameters(parameters);

        var doc = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var page = doc.Pages[p.PageIndex];

        return new GetPdfPageDetailsResult
        {
            PageIndex = p.PageIndex,
            Width = page.Rect.Width,
            Height = page.Rect.Height,
            Rotation = page.Rotate.ToString(),
            MediaBox = new PdfPageBox
            {
                Llx = page.MediaBox.LLX,
                Lly = page.MediaBox.LLY,
                Urx = page.MediaBox.URX,
                Ury = page.MediaBox.URY
            },
            CropBox = new PdfPageBox
            {
                Llx = page.CropBox.LLX,
                Lly = page.CropBox.LLY,
                Urx = page.CropBox.URX,
                Ury = page.CropBox.URY
            },
            Annotations = page.Annotations.Count,
            Paragraphs = page.Paragraphs.Count,
            Images = page.Resources?.Images?.Count ?? 0
        };
    }

    /// <summary>
    ///     Extracts parameters for get_details operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetDetailsParameters ExtractGetDetailsParameters(OperationParameters parameters)
    {
        return new GetDetailsParameters(
            parameters.GetRequired<int>("pageIndex")
        );
    }

    /// <summary>
    ///     Parameters for get_details operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    private sealed record GetDetailsParameters(int PageIndex);
}
