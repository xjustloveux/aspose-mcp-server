using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Handler for adding image stamps to PDF document headers or footers.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddImagePdfHeaderFooterHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_image";

    /// <summary>
    ///     Adds an image stamp to the header or footer of specified pages.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imagePath.
    ///     Optional: position (default: header), alignment (default: center),
    ///     margin (default: 20.0), width (default: 0 = auto), height (default: 0 = auto), pageRange.
    /// </param>
    /// <returns>Success message with the number of pages affected.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the image path is invalid.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file does not exist.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        SecurityHelper.ValidateFilePath(p.ImagePath, "imagePath", true);

        if (!File.Exists(p.ImagePath))
            throw new FileNotFoundException($"Image file not found: {p.ImagePath}");

        var document = context.Document;

        var stamp = new ImageStamp(p.ImagePath);

        if (p.Position.Equals("header", StringComparison.OrdinalIgnoreCase))
        {
            stamp.VerticalAlignment = VerticalAlignment.Top;
            stamp.TopMargin = p.Margin;
        }
        else
        {
            stamp.VerticalAlignment = VerticalAlignment.Bottom;
            stamp.BottomMargin = p.Margin;
        }

        stamp.HorizontalAlignment = AddTextPdfHeaderFooterHandler.ResolveHorizontalAlignment(p.Alignment);

        if (p.Width > 0)
            stamp.Width = p.Width;

        if (p.Height > 0)
            stamp.Height = p.Height;

        var pages = PdfWatermarkHelper.ParsePageRange(p.PageRange, document.Pages.Count);
        foreach (var pageIdx in pages) document.Pages[pageIdx].AddStamp(stamp);

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Added image {p.Position} to {pages.Count} page(s)."
        };
    }

    /// <summary>
    ///     Extracts add image header/footer parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddImageParameters ExtractParameters(OperationParameters parameters)
    {
        return new AddImageParameters(
            parameters.GetRequired<string>("imagePath"),
            parameters.GetOptional("position", "header"),
            parameters.GetOptional("alignment", "center"),
            parameters.GetOptional("margin", 20.0),
            parameters.GetOptional("width", 0.0),
            parameters.GetOptional("height", 0.0),
            parameters.GetOptional<string?>("pageRange")
        );
    }

    /// <summary>
    ///     Parameters for adding an image header or footer.
    /// </summary>
    /// <param name="ImagePath">The path to the image file.</param>
    /// <param name="Position">The position (header or footer).</param>
    /// <param name="Alignment">The horizontal alignment (left, center, or right).</param>
    /// <param name="Margin">The margin from the edge in points.</param>
    /// <param name="Width">The image width (0 for auto).</param>
    /// <param name="Height">The image height (0 for auto).</param>
    /// <param name="PageRange">The optional page range string.</param>
    private sealed record AddImageParameters(
        string ImagePath,
        string Position,
        string Alignment,
        double Margin,
        double Width,
        double Height,
        string? PageRange);
}
