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
    /// <summary>
    ///     Largest scale factor a caller may ask for.
    ///     <para>
    ///         <c>GetThumbnail</c> multiplies the slide's own size by this, so an unbounded value
    ///         is an unbounded bitmap. Twenty times a standard slide is already a very large
    ///         image; beyond that the caller wants a different tool (R4-R08).
    ///     </para>
    /// </summary>
    private const float MaxScale = 20f;

    private const string ScaleParameterName = "scale";

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

        var presentation = context.Document;
        var slideIndexList = PptImageHelper.ParseSlideIndexes(p.SlideIndexes, presentation.Slides.Count);

        // GetThumbnail multiplies the slide's own size by the caller's scale, and nothing bounded
        // either: the structural guard only looks for a Resolution assignment, which this API does
        // not use (R4-R08). Priced before anything is created, so a refusal leaves no directory
        // behind.
        if (!float.IsFinite(p.Scale) || p.Scale <= 0f || p.Scale > MaxScale)
            throw new ArgumentException(
                $"{ScaleParameterName} must be between 0 and {MaxScale} (was {p.Scale}).",
                nameof(parameters));
        RenderBudget.EnsureOutputCount(slideIndexList.Count, "slide images");

        var widthInches = presentation.SlideSize.Size.Width / 72.0 * p.Scale;
        var heightInches = presentation.SlideSize.Size.Height / 72.0 * p.Scale;
        var budget = new PixelBudget();
        foreach (var _ in slideIndexList)
            budget.Add(widthInches, heightInches, 96);

        // Validated before it is created. Creating the caller's directory first meant a request
        // for somewhere outside the allowlist left a directory behind even though it was refused
        // (R7-F03).
        var outputDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputDir,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputDir");
        Directory.CreateDirectory(outputDir);

        // Every slide is staged and the batch published at the end, so an export refused on its
        // last slide leaves the directory as it found it rather than half-replaced (R5-R01).
        using var batch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "slide images",
            context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);

        foreach (var i in slideIndexList)
        {
            using var bmp = presentation.Slides[i].GetThumbnail(p.Scale, p.Scale);
            var fileName = Path.Combine(outputDir, $"slide_{i + 1}.{p.Extension}");
            // H26: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            fileName = SecurityHelper.ResolveAndEnsureWithinAllowlist(fileName,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(fileName));
            // Written through what is left of the request's budget: saving straight to the
            // destination left this path with a file count and no byte limit at all (R4-R08).
            batch.Stage(fileName,
                // ReSharper disable once AccessToDisposedClosure -- Stage invokes its writer synchronously.
                stream => bmp.Save(stream, p.IsJpeg ? ImageFormat.Jpeg : ImageFormat.Png));
        }

        var exportedCount = batch.Publish().Count;

        return new SuccessResult
            { Message = $"Exported {exportedCount} slides. Output: {outputDir}" };
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
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);
        var slideIndexes = parameters.GetOptional<string?>("slideIndexes");
        var formatStr = parameters.GetOptional("format", "png");
        var scale = parameters.GetOptional("scale", 1.0f);

        var isJpeg = formatStr.ToLower() is "jpeg" or "jpg";
        var extension = isJpeg ? "jpg" : "png";

        return new ExportParameters(outputDir, slideIndexes, isJpeg, extension, scale);
    }

    /// <summary>
    ///     Record for holding export slides parameters.
    /// </summary>
    /// <param name="OutputDir">The output directory.</param>
    /// <param name="SlideIndexes">The optional slide indexes string.</param>
    /// <param name="IsJpeg">Whether the output format is JPEG (false = PNG).</param>
    /// <param name="Extension">The file extension.</param>
    /// <param name="Scale">The scale factor.</param>
    private sealed record ExportParameters(
        string OutputDir,
        string? SlideIndexes,
        bool IsJpeg,
        string Extension,
        float Scale);
}
