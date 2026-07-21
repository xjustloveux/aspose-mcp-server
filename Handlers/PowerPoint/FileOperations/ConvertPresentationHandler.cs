using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Progress;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for converting PowerPoint presentations to other formats.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractConvertParameters(parameters);

        var sourcePath = p.InputPath ?? p.Path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required for convert operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
        // H22: resolve symlinks immediately before all sinks (bug 20260415-symlink-toctou-sweep).
        var resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputPath");

        var format = p.Format.ToLower();

        Presentation presentation;
        Presentation? ownedPresentation = null;
        IDisposable? usageScope = null;
        string sourceDescription;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            // Hold a usage scope for the whole conversion so a concurrent close/cleanup cannot
            // dispose the session document mid-operation (same contract as DocumentContext.Create).
            var session = context.SessionManager.GetSession(p.SessionId, identity);
            usageScope = session.AcquireUsage();
            presentation = session.GetDocument<Presentation>();
            sourceDescription = $"session {p.SessionId}";
        }
        else
        {
            SecurityHelper.ValidateFilePath(sourcePath!, "inputPath", true);
            // The input is a read sink: resolve symlinks and enforce the allowlist like the output path.
            var resolvedSourcePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(sourcePath!,
                context.ServerConfig?.AllowedBasePaths ?? [], "inputPath");
            ownedPresentation = new Presentation(resolvedSourcePath);
            presentation = ownedPresentation;
            sourceDescription = sourcePath!;
        }

        try
        {
            if (format is "jpg" or "jpeg" or "png")
            {
                var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

                var slideSize = presentation.SlideSize.Size;
                var targetSize = new Size((int)slideSize.Width, (int)slideSize.Height);

                using var bitmap = slide.GetThumbnail(targetSize);
                var imageFormat = format == "png" ? ImageFormat.Png : ImageFormat.Jpeg;
                bitmap.Save(resolvedOutputPath, imageFormat);

                var formatName = format == "png" ? "PNG" : "JPEG";
                return new SuccessResult
                {
                    Message =
                        $"Slide {p.SlideIndex} from {sourceDescription} converted to {formatName}. Output: {p.OutputPath}"
                };
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

            if (format == "pdf" && context.Progress != null)
            {
                var pdfOptions = new PdfOptions
                {
                    ProgressCallback = new SlidesProgressAdapter(context.Progress)
                };
                presentation.Save(resolvedOutputPath, SaveFormat.Pdf, pdfOptions);
            }
            else
            {
                presentation.Save(resolvedOutputPath, saveFormat);
            }

            return new SuccessResult
            {
                Message =
                    $"Presentation from {sourceDescription} converted to {format.ToUpper()} format. Output: {p.OutputPath}"
            };
        }
        finally
        {
            ownedPresentation?.Dispose();
            usageScope?.Dispose();
        }
    }

    /// <summary>
    ///     Extracts convert parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted convert parameters.</returns>
    private static ConvertParameters ExtractConvertParameters(OperationParameters parameters)
    {
        return new ConvertParameters(
            parameters.GetOptional<string?>("inputPath"),
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetRequired<string>("outputPath"),
            parameters.GetRequired<string>("format"),
            parameters.GetOptional("slideIndex", 0));
    }

    /// <summary>
    ///     Record for holding convert presentation parameters.
    /// </summary>
    /// <param name="InputPath">The input file path.</param>
    /// <param name="Path">Alternative input file path.</param>
    /// <param name="SessionId">The session ID.</param>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="Format">The target format.</param>
    /// <param name="SlideIndex">The slide index for image export.</param>
    private sealed record ConvertParameters(
        string? InputPath,
        string? Path,
        string? SessionId,
        string OutputPath,
        string Format,
        int SlideIndex);
}
