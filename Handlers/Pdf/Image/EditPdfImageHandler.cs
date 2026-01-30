using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ImageFormat = Aspose.Pdf.Drawing.ImageFormat;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for editing images in PDF documents (move or replace).
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditPdfImageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing image on the specified page (move or replace).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex, imageIndex, imagePath, x, y, width, height
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var document = context.Document;

        var actualPageIndex = p.PageIndex < 1 ? 1 : p.PageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (p.ImageIndex < 1 || p.ImageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        string? tempImagePath = null;
        var imagePath = p.ImagePath;
        try
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                tempImagePath = Path.Combine(Path.GetTempPath(), $"temp_image_{Guid.NewGuid()}.png");
                using var imageStream = new FileStream(tempImagePath, FileMode.Create);
                images[p.ImageIndex].Save(imageStream, ImageFormat.Png);
                imagePath = tempImagePath;
            }
            else
            {
                SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            images.Delete(p.ImageIndex);
            var newX = p.X ?? 100;
            var newY = p.Y ?? 600;
            page.AddImage(imagePath,
                new Rectangle(newX, newY, p.Width.HasValue ? newX + p.Width.Value : newX + 200,
                    p.Height.HasValue ? newY + p.Height.Value : newY + 200));

            MarkModified(context);

            var action = tempImagePath != null ? "Moved" : "Replaced";
            return new SuccessResult { Message = $"{action} image {p.ImageIndex} on page {p.PageIndex}." };
        }
        finally
        {
            if (tempImagePath != null && File.Exists(tempImagePath))
                File.Delete(tempImagePath);
        }
    }

    /// <summary>
    ///     Extracts edit parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("imageIndex", 1),
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional<double?>("x"),
            parameters.GetOptional<double?>("y"),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"));
    }

    /// <summary>
    ///     Parameters for editing an image.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="ImageIndex">The 1-based image index.</param>
    /// <param name="ImagePath">The optional path to the replacement image.</param>
    /// <param name="X">The optional X coordinate.</param>
    /// <param name="Y">The optional Y coordinate.</param>
    /// <param name="Width">The optional width.</param>
    /// <param name="Height">The optional height.</param>
    private sealed record EditParameters(
        int PageIndex,
        int ImageIndex,
        string? ImagePath,
        double? X,
        double? Y,
        double? Width,
        double? Height);
}
