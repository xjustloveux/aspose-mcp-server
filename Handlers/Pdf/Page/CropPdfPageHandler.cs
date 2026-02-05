using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for cropping pages in PDF documents by adjusting the crop box.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CropPdfPageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "crop";

    /// <summary>
    ///     Crops a page by setting the CropBox rectangle.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based), x (lower-left X), y (lower-left Y), width, height
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractCropParameters(parameters);

        var doc = context.Document;
        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex {p.PageIndex} is out of range (1-{doc.Pages.Count})");

        var page = doc.Pages[p.PageIndex];
        var cropBox = new Rectangle(p.X, p.Y, p.X + p.Width, p.Y + p.Height);
        page.CropBox = cropBox;

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Page {p.PageIndex} cropped to ({p.X}, {p.Y}, {p.Width}x{p.Height})."
        };
    }

    /// <summary>
    ///     Extracts crop parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted crop parameters.</returns>
    private static CropParameters ExtractCropParameters(OperationParameters parameters)
    {
        return new CropParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<double>("x"),
            parameters.GetRequired<double>("y"),
            parameters.GetRequired<double>("width"),
            parameters.GetRequired<double>("height")
        );
    }

    /// <summary>
    ///     Parameters for crop operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="X">The lower-left X coordinate.</param>
    /// <param name="Y">The lower-left Y coordinate.</param>
    /// <param name="Width">The crop width.</param>
    /// <param name="Height">The crop height.</param>
    private sealed record CropParameters(int PageIndex, double X, double Y, double Width, double Height);
}
