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
        var imagePath = parameters.GetRequired<string>("imagePath");
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 600.0);
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var document = context.Document;

        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        page.AddImage(imagePath,
            new Rectangle(x, y, width.HasValue ? x + width.Value : x + 200,
                height.HasValue ? y + height.Value : y + 200));

        MarkModified(context);

        return Success($"Added image to page {actualPageIndex}.");
    }
}
