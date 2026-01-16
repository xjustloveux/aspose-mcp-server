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
    /// <exception cref="ArgumentException">Thrown when imagePath is missing or image cannot be decoded.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddImageWatermarkParameters(parameters);

        if (string.IsNullOrEmpty(p.ImagePath))
            throw new ArgumentException("imagePath is required for add_image operation");

        SecurityHelper.ValidateFilePath(p.ImagePath, "imagePath", true);

        if (!System.IO.File.Exists(p.ImagePath))
            throw new FileNotFoundException($"Image file not found: {p.ImagePath}");

        var doc = context.Document;

        var watermarkOptions = new ImageWatermarkOptions
        {
            Scale = p.Scale,
            IsWashout = p.IsWashout
        };

        using var bitmap = SKBitmap.Decode(p.ImagePath);
        if (bitmap == null)
            throw new ArgumentException(
                $"Failed to decode image: {p.ImagePath}. Ensure the file is a valid image format.");

        doc.Watermark.SetImage(bitmap, watermarkOptions);

        MarkModified(context);

        return Success(
            $"Image watermark added to document. Image: {Path.GetFileName(p.ImagePath)}, Scale: {p.Scale}, Washout: {p.IsWashout}");
    }

    private static AddImageWatermarkParameters ExtractAddImageWatermarkParameters(OperationParameters parameters)
    {
        return new AddImageWatermarkParameters(
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("scale", 1.0),
            parameters.GetOptional("isWashout", true));
    }

    private sealed record AddImageWatermarkParameters(
        string? ImagePath,
        double Scale,
        bool IsWashout);
}
