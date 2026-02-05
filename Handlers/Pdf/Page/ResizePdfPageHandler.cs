using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for resizing pages in PDF documents by adjusting the MediaBox.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ResizePdfPageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "resize";

    /// <summary>
    ///     Resizes a page by setting the MediaBox dimensions.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex (1-based), width, height
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractResizeParameters(parameters);

        var doc = context.Document;
        if (p.PageIndex < 1 || p.PageIndex > doc.Pages.Count)
            throw new ArgumentException($"pageIndex {p.PageIndex} is out of range (1-{doc.Pages.Count})");

        if (p.Width <= 0 || p.Height <= 0)
            throw new ArgumentException("width and height must be positive values");

        var page = doc.Pages[p.PageIndex];
        page.MediaBox = new Rectangle(0, 0, p.Width, p.Height);

        // Also update CropBox to match MediaBox to ensure consistent rendering
        page.CropBox = new Rectangle(0, 0, p.Width, p.Height);

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Page {p.PageIndex} resized to {p.Width}x{p.Height} points."
        };
    }

    /// <summary>
    ///     Extracts resize parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted resize parameters.</returns>
    private static ResizeParameters ExtractResizeParameters(OperationParameters parameters)
    {
        return new ResizeParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<double>("width"),
            parameters.GetRequired<double>("height")
        );
    }

    /// <summary>
    ///     Parameters for resize operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="Width">The new page width in points.</param>
    /// <param name="Height">The new page height in points.</param>
    private sealed record ResizeParameters(int PageIndex, double Width, double Height);
}
