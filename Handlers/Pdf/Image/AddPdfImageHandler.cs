using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for adding images to PDF documents.
/// </summary>
public class AddPdfImageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an image to the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imagePath
    ///     Optional: pageIndex (default: 1), x (default: 100), y (default: 600), width, height
    /// </param>
    /// <returns>Success message with add details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        SecurityHelper.ValidateFilePath(p.ImagePath, "imagePath", true);

        if (!File.Exists(p.ImagePath))
            throw new FileNotFoundException($"Image file not found: {p.ImagePath}");

        var document = context.Document;

        var actualPageIndex = p.PageIndex < 1 ? 1 : p.PageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        page.AddImage(p.ImagePath,
            new Rectangle(p.X, p.Y, p.Width.HasValue ? p.X + p.Width.Value : p.X + 200,
                p.Height.HasValue ? p.Y + p.Height.Value : p.Y + 200));

        MarkModified(context);

        return Success($"Added image to page {actualPageIndex}.");
    }

    /// <summary>
    ///     Extracts add parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<string>("imagePath"),
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 600.0),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"));
    }

    /// <summary>
    ///     Parameters for adding an image.
    /// </summary>
    /// <param name="ImagePath">The path to the image file.</param>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The optional width.</param>
    /// <param name="Height">The optional height.</param>
    private record AddParameters(
        string ImagePath,
        int PageIndex,
        double X,
        double Y,
        double? Width,
        double? Height);
}
