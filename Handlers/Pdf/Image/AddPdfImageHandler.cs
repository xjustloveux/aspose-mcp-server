using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for adding images to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        SecurityHelper.ValidateFilePath(p.ImagePath, "imagePath", true);
        var resolvedImagePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.ImagePath,
            context.ServerConfig?.AllowedBasePaths ?? [], "imagePath");

        if (!File.Exists(resolvedImagePath))
            throw new FileNotFoundException("The specified file was not found.");

        var document = context.Document;

        // Reject non-positive indices instead of clamping to page 1: 'get' treats pageIndex 0
        // as "all pages", so a silent clamp would mutate a page the caller never named.
        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];
        page.AddImage(resolvedImagePath,
            new Rectangle(p.X, p.Y, p.Width.HasValue ? p.X + p.Width.Value : p.X + 200,
                p.Height.HasValue ? p.Y + p.Height.Value : p.Y + 200));

        MarkModified(context);

        return new SuccessResult { Message = $"Added image to page {p.PageIndex}." };
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
    private sealed record AddParameters(
        string ImagePath,
        int PageIndex,
        double X,
        double Y,
        double? Width,
        double? Height);
}
