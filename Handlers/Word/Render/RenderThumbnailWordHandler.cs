using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Render;

namespace AsposeMcpServer.Handlers.Word.Render;

/// <summary>
///     Handler for rendering a thumbnail of the first page of a Word document.
/// </summary>
[ResultType(typeof(RenderResult))]
public class RenderThumbnailWordHandler : OperationHandlerBase<Document>
{
    private const string OutputPathParameter = "outputPath";

    /// <inheritdoc />
    public override string Operation => "thumbnail";

    /// <summary>
    ///     Renders a thumbnail of the first page of a Word document.
    ///     The thumbnail is a lower-resolution image suitable for previews.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (output image file path)
    ///     Optional: format (png, jpeg; default: png), scale (default: 0.25)
    /// </param>
    /// <returns>Render result with output file path.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or scale is invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractThumbnailParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.Path,
            context.ServerConfig?.AllowedBasePaths ?? [], "path");
        SecurityHelper.ValidateFilePath(p.OutputPath, OutputPathParameter, true);
        SecurityHelper.ValidateNumericRange((long)(p.Scale * 1000), "scale", 1, 10_000);
        // Refused early; the sink below resolves again immediately before it writes.
        SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);

        if (p.Scale <= 0 || p.Scale > 1)
            throw new ArgumentException("scale must be between 0 (exclusive) and 1 (inclusive)");

        var doc = GuardedWordLoader.Load(resolvedPath, context.ServerConfig?.AllowedBasePaths ?? []);

        var saveFormat = p.Format.ToLowerInvariant() switch
        {
            "png" => SaveFormat.Png,
            "jpeg" or "jpg" => SaveFormat.Jpeg,
            _ => throw new ArgumentException(
                $"Unknown thumbnail format: {p.Format}. Supported: png, jpeg")
        };

        // One page at a fraction of 96 DPI is small, but the page itself can be large, so the
        // request is priced from the real page size rather than assumed to be cheap (R2-R01).
        // The measurement comes from the page that will actually be drawn (R3-R03).
        var pageInfo = doc.PageCount > 0 ? doc.GetPageInfo(0) : null;
        new PixelBudget().Add((pageInfo?.WidthInPoints ?? 0) / 72.0,
            (pageInfo?.HeightInPoints ?? 0) / 72.0, (int)Math.Ceiling(96 * p.Scale));

        var options = new ImageSaveOptions(saveFormat)
        {
            Resolution = (float)(96 * p.Scale),
            PageSet = new PageSet(0)
        };

        // Creating a directory is itself a filesystem write, so it is derived from the resolved
        // path rather than the caller's string (R7-T01).
        var outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        // H1: resolve symlinks immediately before the sink to close TOCTOU (bug 20260415-symlink-toctou-sweep).
        outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
        doc.Save(outputPath, options);

        return new RenderResult
        {
            Message = $"Thumbnail rendered at {p.Scale:P0} scale in {p.Format.ToUpperInvariant()} format.",
            OutputPaths = [p.OutputPath],
            Format = p.Format
        };
    }

    /// <summary>
    ///     Extracts thumbnail parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted thumbnail parameters.</returns>
    private static ThumbnailParameters ExtractThumbnailParameters(OperationParameters parameters)
    {
        return new ThumbnailParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>(OutputPathParameter),
            parameters.GetOptional("format", "png"),
            parameters.GetOptional("scale", 0.25)
        );
    }

    /// <summary>
    ///     Parameters for the thumbnail operation.
    /// </summary>
    /// <param name="Path">The source document file path.</param>
    /// <param name="OutputPath">The output image file path.</param>
    /// <param name="Format">The output image format.</param>
    /// <param name="Scale">The scale factor (0-1).</param>
    private sealed record ThumbnailParameters(
        string Path,
        string OutputPath,
        string Format,
        double Scale);
}
