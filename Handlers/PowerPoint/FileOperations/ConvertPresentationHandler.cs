using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for converting PowerPoint presentations to other formats.
/// </summary>
public class ConvertPresentationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "convert";

    /// <summary>
    ///     Converts a PowerPoint presentation to another format.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, format
    ///     Optional: inputPath, path, sessionId, slideIndex
    /// </param>
    /// <returns>Success message with output path.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var inputPath = parameters.GetOptional<string?>("inputPath");
        var path = parameters.GetOptional<string?>("path");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputPath = parameters.GetRequired<string>("outputPath");
        var format = parameters.GetRequired<string>("format");
        var slideIndex = parameters.GetOptional("slideIndex", 0);

        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required for convert operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        format = format.ToLower();

        Presentation presentation;
        string sourceDescription;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            presentation = context.SessionManager.GetDocument<Presentation>(sessionId, identity);
            sourceDescription = $"session {sessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(sourcePath!, "inputPath", true);
            presentation = new Presentation(sourcePath);
            sourceDescription = sourcePath!;
        }

        if (format is "jpg" or "jpeg" or "png")
        {
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var slideSize = presentation.SlideSize.Size;
            var targetSize = new Size((int)slideSize.Width, (int)slideSize.Height);

#pragma warning disable CA1416
            using var bitmap = slide.GetThumbnail(targetSize);
            var imageFormat = format == "png" ? ImageFormat.Png : ImageFormat.Jpeg;
            bitmap.Save(outputPath, imageFormat);
#pragma warning restore CA1416

            var formatName = format == "png" ? "PNG" : "JPEG";
            return Success(
                $"Slide {slideIndex} from {sourceDescription} converted to {formatName}. Output: {outputPath}");
        }

        var saveFormat = format switch
        {
            "pdf" => SaveFormat.Pdf,
            "html" => SaveFormat.Html,
            "pptx" => SaveFormat.Pptx,
            "ppt" => SaveFormat.Ppt,
            "odp" => SaveFormat.Odp,
            "xps" => SaveFormat.Xps,
            "tiff" => SaveFormat.Tiff,
            _ => throw new ArgumentException($"Unsupported format: {format}")
        };

        presentation.Save(outputPath, saveFormat);

        return Success(
            $"Presentation from {sourceDescription} converted to {format.ToUpper()} format. Output: {outputPath}");
    }
}
