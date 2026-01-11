using System.Drawing.Imaging;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for editing images in PDF documents (move or replace).
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var imageIndex = parameters.GetOptional("imageIndex", 1);
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var x = parameters.GetOptional<double?>("x");
        var y = parameters.GetOptional<double?>("y");
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");

        var document = context.Document;

        var actualPageIndex = pageIndex < 1 ? 1 : pageIndex;
        if (actualPageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[actualPageIndex];
        var images = page.Resources?.Images;
        if (images == null)
            throw new ArgumentException("No images found on the page");
        if (imageIndex < 1 || imageIndex > images.Count)
            throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

        string? tempImagePath = null;
        try
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                tempImagePath = Path.Combine(Path.GetTempPath(), $"temp_image_{Guid.NewGuid()}.png");
                using var imageStream = new FileStream(tempImagePath, FileMode.Create);
#pragma warning disable CA1416
                images[imageIndex].Save(imageStream, ImageFormat.Png);
#pragma warning restore CA1416
                imagePath = tempImagePath;
            }
            else
            {
                SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Image file not found: {imagePath}");
            }

            images.Delete(imageIndex);
            var newX = x ?? 100;
            var newY = y ?? 600;
            page.AddImage(imagePath,
                new Rectangle(newX, newY, width.HasValue ? newX + width.Value : newX + 200,
                    height.HasValue ? newY + height.Value : newY + 200));

            MarkModified(context);

            var action = tempImagePath != null ? "Moved" : "Replaced";
            return Success($"{action} image {imageIndex} on page {pageIndex}.");
        }
        finally
        {
            if (tempImagePath != null && File.Exists(tempImagePath))
                File.Delete(tempImagePath);
        }
    }
}
