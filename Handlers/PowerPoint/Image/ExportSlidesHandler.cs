using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for exporting PowerPoint slides as images.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = context.SourcePath;
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for export_slides operation");

        SecurityHelper.ValidateFilePath(path, "path", true);

        var p = ExtractExportParameters(parameters, path);

        Directory.CreateDirectory(p.OutputDir);

        var presentation = context.Document;
        var slideIndexList = PptImageHelper.ParseSlideIndexes(p.SlideIndexes, presentation.Slides.Count);

        var exportedCount = 0;
        foreach (var i in slideIndexList)
        {
            using var bmp = presentation.Slides[i].GetThumbnail(p.Scale, p.Scale);
            var fileName = Path.Combine(p.OutputDir, $"slide_{i + 1}.{p.Extension}");
            // CA1416 - System.Drawing.Common is Windows-only, cross-platform support not required
#pragma warning disable CA1416
            bmp.Save(fileName, p.Format);
#pragma warning restore CA1416
            exportedCount++;
        }

        return new SuccessResult
            { Message = $"Exported {exportedCount} slides. Output: {Path.GetFullPath(p.OutputDir)}" };
    }

    /// <summary>
    ///     Extracts export parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <param name="path">The source file path.</param>
    /// <returns>The extracted export parameters.</returns>
    private static ExportParameters ExtractExportParameters(OperationParameters parameters, string path)
    {
        var outputDir = parameters.GetOptional<string?>("outputDir") ?? Path.GetDirectoryName(path) ?? ".";
        var slideIndexes = parameters.GetOptional<string?>("slideIndexes");
        var formatStr = parameters.GetOptional("format", "png");
        var scale = parameters.GetOptional("scale", 1.0f);

        // CA1416 - System.Drawing.Common is Windows-only, cross-platform support not required
#pragma warning disable CA1416
        var format = formatStr.ToLower() switch
        {
            "jpeg" or "jpg" => ImageFormat.Jpeg,
            _ => ImageFormat.Png
        };
        var extension = format.Equals(ImageFormat.Png) ? "png" : "jpg";
#pragma warning restore CA1416

        return new ExportParameters(outputDir, slideIndexes, format, extension, scale);
    }

    /// <summary>
    ///     Record for holding export slides parameters.
    /// </summary>
    /// <param name="OutputDir">The output directory.</param>
    /// <param name="SlideIndexes">The optional slide indexes string.</param>
    /// <param name="Format">The image format.</param>
    /// <param name="Extension">The file extension.</param>
    /// <param name="Scale">The scale factor.</param>
    private sealed record ExportParameters(
        string OutputDir,
        string? SlideIndexes,
        ImageFormat Format,
        string Extension,
        float Scale);
}
