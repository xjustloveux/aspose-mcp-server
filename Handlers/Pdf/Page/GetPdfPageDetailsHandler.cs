using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for getting detailed information about a specific page in PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetDetailsParameters(parameters);

        var doc = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {doc.Pages.Count}");

        var page = doc.Pages[p.PageIndex];
        var result = new
        {
            pageIndex = p.PageIndex,
            width = page.Rect.Width,
            height = page.Rect.Height,
            rotation = page.Rotate.ToString(),
            mediaBox = new
            {
                llx = page.MediaBox.LLX,
                lly = page.MediaBox.LLY,
                urx = page.MediaBox.URX,
                ury = page.MediaBox.URY
            },
            cropBox = new
            {
                llx = page.CropBox.LLX,
                lly = page.CropBox.LLY,
                urx = page.CropBox.URX,
                ury = page.CropBox.URY
            },
            annotations = page.Annotations.Count,
            paragraphs = page.Paragraphs.Count,
            images = page.Resources?.Images?.Count ?? 0
        };

        return JsonResult(result);
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
