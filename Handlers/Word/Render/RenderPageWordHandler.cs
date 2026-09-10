using Aspose.Words;
using Aspose.Words.Saving;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Render;

namespace AsposeMcpServer.Handlers.Word.Render;

/// <summary>
///     Handler for rendering specific pages of a Word document to images.
/// </summary>
[ResultType(typeof(RenderResult))]
public class RenderPageWordHandler : OperationHandlerBase<Document>
{
    private const string OutputPathParameter = "outputPath";

    /// <inheritdoc />
    public override string Operation => "render";

    /// <summary>
    ///     Renders a specific page (or all pages) of a Word document to image file(s).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (source file path), outputPath (output file/directory path)
    ///     Optional: pageIndex (1-based, default: all pages), format (png, jpeg, bmp, tiff, svg; default: png), dpi (default:
    ///     150)
    /// </param>
    /// <returns>Render result with output file paths.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractRenderParameters(parameters);

        SecurityHelper.ValidateFilePath(p.Path, allowAbsolutePaths: true);
        var resolvedPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.Path,
            context.ServerConfig?.AllowedBasePaths ?? [], "path");
        SecurityHelper.ValidateFilePath(p.OutputPath, OutputPathParameter, true);
        SecurityHelper.ValidateNumericRange(p.Dpi, "dpi", 10, 1200);
        _ = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);

        var doc = GuardedWordLoader.Load(resolvedPath, context.ServerConfig?.AllowedBasePaths ?? []);

        // Price the whole request before rendering anything. Each parameter is individually
        // capped, but the cost is their product, and the real page size matters: a document set
        // to a large custom page costs several times an A4 page at the same DPI (R2-R01). Every
        // page is measured on its own, because the first section's setup says nothing about a
        // later section that switches to a much larger sheet (R3-R03).
        RenderBudget.EnsureOutputCount(p.PageIndex.HasValue ? 1 : doc.PageCount, "image files");

        var budget = new PixelBudget();
        if (p.PageIndex.HasValue)
        {
            if (p.PageIndex.Value >= 1 && p.PageIndex.Value <= doc.PageCount)
                AddPageToBudget(budget, doc, p.PageIndex.Value - 1, p.Dpi);
        }
        else
        {
            for (var page = 0; page < doc.PageCount; page++)
                AddPageToBudget(budget, doc, page, p.Dpi);
        }

        var saveFormat = ResolveSaveFormat(p.Format);
        var outputPaths = new List<string>();

        if (p.PageIndex.HasValue)
        {
            if (p.PageIndex.Value < 1 || p.PageIndex.Value > doc.PageCount)
                throw new ArgumentException(
                    $"pageIndex {p.PageIndex.Value} is out of range (1-{doc.PageCount})");

            var options = CreateImageSaveOptions(saveFormat, p.Dpi, p.PageIndex.Value - 1);

            // Creating a directory is itself a filesystem write, so it is derived from the
            // resolved path: taking it from the caller's string left a directory behind at a
            // location the allowlist was about to refuse (R7-T01).
            var resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
            var outputDir = Path.GetDirectoryName(resolvedOutputPath);
            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            // H2: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
            resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(resolvedOutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
            doc.Save(resolvedOutputPath, options);
            outputPaths.Add(resolvedOutputPath);
        }
        else
        {
            // The per-page names are built from the resolved template for the same reason
            // (R7-T01).
            var resolvedTemplate = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
            var outputDir = Path.GetDirectoryName(resolvedTemplate);
            var baseName = Path.GetFileNameWithoutExtension(resolvedTemplate);
            var ext = Path.GetExtension(resolvedTemplate);
            if (string.IsNullOrEmpty(ext)) ext = $".{p.Format}";

            if (!string.IsNullOrEmpty(outputDir))
                Directory.CreateDirectory(outputDir);

            for (var i = 0; i < doc.PageCount; i++)
            {
                var options = CreateImageSaveOptions(saveFormat, p.Dpi, i);
                var pagePath = Path.Combine(outputDir ?? ".",
                    $"{baseName}_page_{i + 1}{ext}");
                // H2: resolve symlinks immediately before each per-page sink (bug 20260415-symlink-toctou-sweep).
                pagePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(pagePath,
                    context.ServerConfig?.AllowedBasePaths ?? [], nameof(pagePath));
                doc.Save(pagePath, options);
                outputPaths.Add(pagePath);
            }
        }

        var message = p.PageIndex.HasValue
            ? $"Page {p.PageIndex.Value} rendered to {p.Format.ToUpperInvariant()} format."
            : $"{outputPaths.Count} page(s) rendered to {p.Format.ToUpperInvariant()} format.";

        return new RenderResult
        {
            Message = message,
            OutputPaths = outputPaths,
            Format = p.Format
        };
    }

    /// <summary>
    ///     Resolves the save format from a format string.
    /// </summary>
    /// <param name="format">The format string.</param>
    /// <returns>The corresponding SaveFormat value.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is unknown.</exception>
    private static SaveFormat ResolveSaveFormat(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "png" => SaveFormat.Png,
            "jpeg" or "jpg" => SaveFormat.Jpeg,
            "bmp" => SaveFormat.Bmp,
            "tiff" or "tif" => SaveFormat.Tiff,
            "svg" => SaveFormat.Svg,
            _ => throw new ArgumentException(
                $"Unknown render format: {format}. Supported: png, jpeg, bmp, tiff, svg")
        };
    }

    /// <summary>
    ///     Creates ImageSaveOptions for a specific page.
    /// </summary>
    /// <param name="saveFormat">The save format.</param>
    /// <param name="dpi">The rendering DPI.</param>
    /// <param name="pageIndex">The 0-based page index.</param>
    /// <returns>Configured ImageSaveOptions.</returns>
    private static ImageSaveOptions CreateImageSaveOptions(SaveFormat saveFormat, int dpi, int pageIndex)
    {
        return new ImageSaveOptions(saveFormat)
        {
            Resolution = dpi,
            PageSet = new PageSet(pageIndex)
        };
    }

    /// <summary>
    ///     Extracts render parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted render parameters.</returns>
    private static RenderParameters ExtractRenderParameters(OperationParameters parameters)
    {
        return new RenderParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>(OutputPathParameter),
            parameters.GetOptional<int?>("pageIndex"),
            parameters.GetOptional("format", "png"),
            parameters.GetOptional("dpi", 150)
        );
    }

    /// <summary>
    ///     Adds one page's real size to the render budget.
    /// </summary>
    /// <param name="budget">The budget to charge.</param>
    /// <param name="document">The document being rendered.</param>
    /// <param name="pageIndex">Zero-based index of the page about to be rendered.</param>
    /// <param name="dpi">Resolution the page will be rendered at.</param>
    /// <exception cref="ArgumentException">Thrown when the page takes the render past the budget.</exception>
    private static void AddPageToBudget(PixelBudget budget, Document document, int pageIndex, int dpi)
    {
        var info = document.GetPageInfo(pageIndex);
        budget.Add(info.WidthInPoints / 72.0, info.HeightInPoints / 72.0, dpi);
    }

    /// <summary>
    ///     Parameters for the render operation.
    /// </summary>
    /// <param name="Path">The source document file path.</param>
    /// <param name="OutputPath">The output file/directory path.</param>
    /// <param name="PageIndex">The 1-based page index (null for all pages).</param>
    /// <param name="Format">The output image format.</param>
    /// <param name="Dpi">The rendering DPI.</param>
    private sealed record RenderParameters(
        string Path,
        string OutputPath,
        int? PageIndex,
        string Format,
        int Dpi);
}
