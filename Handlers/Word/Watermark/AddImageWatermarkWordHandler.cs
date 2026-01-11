using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using SkiaSharp;

namespace AsposeMcpServer.Handlers.Word.Watermark;

/// <summary>
///     Handler for adding image watermarks to Word documents.
/// </summary>
public class AddImageWatermarkWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_image";

    /// <summary>
    ///     Adds an image watermark to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imagePath
    ///     Optional: scale (default: 1.0), isWashout (default: true)
    /// </param>
    /// <returns>Success message with watermark details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var scale = parameters.GetOptional("scale", 1.0);
        var isWashout = parameters.GetOptional("isWashout", true);

        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add_image operation");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!System.IO.File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = context.Document;

        var watermarkOptions = new ImageWatermarkOptions
        {
            Scale = scale,
            IsWashout = isWashout
        };

        using var bitmap = SKBitmap.Decode(imagePath);
        if (bitmap == null)
            throw new ArgumentException(
                $"Failed to decode image: {imagePath}. Ensure the file is a valid image format.");

        doc.Watermark.SetImage(bitmap, watermarkOptions);

        MarkModified(context);

        return Success(
            $"Image watermark added to document. Image: {Path.GetFileName(imagePath)}, Scale: {scale}, Washout: {isWashout}");
    }
}
