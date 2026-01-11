using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for exporting PowerPoint slides as images.
/// </summary>
public class ExportSlidesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "export_slides";

    /// <summary>
    ///     Exports slides as image files.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: (none, uses context.SourcePath)
    ///     Optional: outputDir, slideIndexes, format, scale
    /// </param>
    /// <returns>Success message with export details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = context.SourcePath;
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for export_slides operation");

        SecurityHelper.ValidateFilePath(path, "path", true);

        var outputDir = parameters.GetOptional<string?>("outputDir");
        var slideIndexesStr = parameters.GetOptional<string?>("slideIndexes");
        var formatStr = parameters.GetOptional("format", "png");
        var scale = parameters.GetOptional("scale", 1.0f);

        var actualOutputDir = outputDir ?? Path.GetDirectoryName(path) ?? ".";

#pragma warning disable CA1416
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416

        Directory.CreateDirectory(actualOutputDir);

        var presentation = context.Document;
        var slideIndexList = PptImageHelper.ParseSlideIndexes(slideIndexesStr, presentation.Slides.Count);

        var exportedCount = 0;
        foreach (var i in slideIndexList)
        {
            using var bmp = presentation.Slides[i].GetThumbnail(scale, scale);
            var fileName = Path.Combine(actualOutputDir, $"slide_{i + 1}.{extension}");
#pragma warning disable CA1416
            bmp.Save(fileName, format);
#pragma warning restore CA1416
            exportedCount++;
        }

        return Success($"Exported {exportedCount} slides. Output: {Path.GetFullPath(actualOutputDir)}");
    }
}
