using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ImageFormat = Aspose.Pdf.Drawing.ImageFormat;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Handler for extracting images from PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractPdfImageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>The most image files one request may produce; the render budget's, unless a fixture lowers it.</summary>
    internal static int OutputFileBudget { get; set; } = RenderBudget.MaxOutputFiles;

    /// <summary>The most bytes one request may produce across all its images; the render budget's, unless a fixture lowers it.</summary>
    internal static long OutputByteBudget { get; set; } = RenderBudget.MaxOutputBytes;

    /// <summary>
    ///     Extracts images from the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex, imageIndex, outputPath, outputDir
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractExtractParameters(parameters);

        // The canonical paths are what the rest of this method uses. Validating the caller's
        // string and then building filenames from that same still-mutable string is the
        // check-to-use gap the resolver exists to close (R7-F03).
        string? outputPath = null;
        if (!string.IsNullOrEmpty(p.OutputPath))
        {
            SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);
            outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], "outputPath");
        }

        string? outputDir = null;
        if (!string.IsNullOrEmpty(p.OutputDir))
        {
            SecurityHelper.ValidateFilePath(p.OutputDir, "outputDir", true);
            outputDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputDir,
                context.ServerConfig?.AllowedBasePaths ?? [], "outputDir");
        }

        // The source's directory comes from the context, which nothing here resolved: it is
        // resolved against the allowlist before a directory is created at it (R22-TST01).
        var targetDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(
            outputDir ?? Path.GetDirectoryName(outputPath) ?? Path.GetDirectoryName(context.SourcePath) ?? ".",
            context.ServerConfig?.AllowedBasePaths ?? [], "targetDir");
        Directory.CreateDirectory(targetDir);

        var document = context.Document;
        // Reject non-positive indices instead of clamping to page 1: 'get' treats pageIndex 0
        // as "all pages", so a silent clamp would read a page the caller never named.
        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];
        var images = page.Resources?.Images;
        if (images == null || images.Count == 0)
            return new SuccessResult { Message = $"No images found on page {p.PageIndex}." };

        if (p.ImageIndex is > 0)
        {
            if (p.ImageIndex.Value < 1 || p.ImageIndex.Value > images.Count)
                throw new ArgumentException($"imageIndex must be between 1 and {images.Count}");

            var image = images[p.ImageIndex.Value];
            var fileName = outputPath ??
                           Path.Combine(targetDir, $"page_{p.PageIndex}_image_{p.ImageIndex.Value}.png");
            // H32: resolve symlinks immediately before the FileStream sink (bug 20260415-symlink-toctou-sweep).
            fileName = SecurityHelper.ResolveAndEnsureWithinAllowlist(fileName,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(fileName));
            // Through the bounded publisher: an unbounded FileStream truncated the caller's
            // destination before the image was known to fit, and left a partial file when it did
            // not (R8-C02).
            BoundedFilePublisher.Publish(fileName, RenderBudget.MaxOutputBytes,
                stream => image.Save(stream, ImageFormat.Png), "extracted image",
                context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);
            return new SuccessResult
                { Message = $"Extracted image {p.ImageIndex.Value} from page {p.PageIndex} to: {fileName}" };
        }

        // The whole request in one batch with one budget: a publisher per image gave every image
        // the full output budget and the request as many images as the page held, and a failure
        // part-way left the earlier images published (R23-PDF01). Count admitted before the
        // first write; bytes charged across all of them; over either, nothing is published.
        if (images.Count > OutputFileBudget)
            throw new ArgumentException(
                $"The request would produce {images.Count:N0} image files, above the limit of "
                + $"{OutputFileBudget:N0}. Extract a subset instead of the whole page.");

        using var batch = new BoundedFileBatch(OutputByteBudget, "extracted images",
            context.Recovery, context.ServerConfig?.AllowedBasePaths ?? []);
        for (var i = 1; i <= images.Count; i++)
        {
            var image = images[i];
            var fileName = Path.Combine(targetDir, $"page_{p.PageIndex}_image_{i}.png");
            // H32: resolve symlinks immediately before the FileStream sink (bug 20260415-symlink-toctou-sweep).
            fileName = SecurityHelper.ResolveAndEnsureWithinAllowlist(fileName,
                context.ServerConfig?.AllowedBasePaths ?? [], nameof(fileName));
            batch.Stage(fileName, stream => image.Save(stream, ImageFormat.Png));
        }

        var published = batch.Publish();

        return new SuccessResult
            { Message = $"Extracted {published.Count} image(s) from page {p.PageIndex} to: {targetDir}" };
    }

    /// <summary>
    ///     Extracts extract parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static ExtractParameters ExtractExtractParameters(OperationParameters parameters)
    {
        return new ExtractParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional<int?>("imageIndex"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional<string?>("outputDir"));
    }

    /// <summary>
    ///     Parameters for extracting images.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="ImageIndex">The optional 1-based image index.</param>
    /// <param name="OutputPath">The optional output file path.</param>
    /// <param name="OutputDir">The optional output directory.</param>
    private sealed record ExtractParameters(int PageIndex, int? ImageIndex, string? OutputPath, string? OutputDir);
}
