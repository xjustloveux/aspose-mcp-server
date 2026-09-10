using System.Drawing.Imaging;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Represents extraction statistics.
/// </summary>
/// <param name="Count">The number of images extracted.</param>
/// <param name="SkippedCount">The number of images skipped.</param>
internal record ExtractionStats(int Count, int SkippedCount);

/// <summary>
///     Handler for extracting embedded images from PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ExtractPptImageHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "extract";

    /// <summary>
    ///     Extracts embedded images from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: (none, uses context.SourcePath)
    ///     Optional: outputDir, format, skipDuplicates
    /// </param>
    /// <returns>Success message with extraction details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = ValidateSourcePath(context.SourcePath);
        var extractParams = ExtractImageParameters(parameters, path);

        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];
        // Resolve before creating: a refused destination must not be brought into existence.
        var resolvedOutputDir = SecurityHelper.ResolveAndEnsureWithinAllowlist(
            extractParams.OutputDir, allowedBasePaths, "outputDir");
        Directory.CreateDirectory(resolvedOutputDir);

        var stats = ExtractAllImages(context.Document, extractParams, allowedBasePaths,
            context.Recovery);

        return new SuccessResult { Message = BuildResultMessage(stats.Count, stats.SkippedCount, extractParams) };
    }

    /// <summary>
    ///     Validates the source path.
    /// </summary>
    /// <param name="path">The source path to validate.</param>
    /// <returns>The validated path.</returns>
    private static string ValidateSourcePath(string? path)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for extract operation");
        SecurityHelper.ValidateFilePath(path, nameof(path), true);
        return path;
    }

    /// <summary>
    ///     Extracts image parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <param name="path">The source file path.</param>
    /// <returns>The extraction parameters.</returns>
    private static ExtractionParameters ExtractImageParameters(OperationParameters parameters, string path)
    {
        var outputDir = parameters.GetOptional<string?>("outputDir") ?? Path.GetDirectoryName(path) ?? ".";
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);
        var formatStr = parameters.GetOptional("format", "png");
        var skipDuplicates = parameters.GetOptional("skipDuplicates", false);

        var isJpeg = formatStr.ToLower() is "jpeg" or "jpg";
        var extension = isJpeg ? "jpg" : "png";

        return new ExtractionParameters(outputDir, isJpeg, extension, skipDuplicates);
    }

    /// <summary>
    ///     Extracts all images from the presentation.
    /// </summary>
    /// <param name="presentation">The presentation to extract images from.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <param name="allowedBasePaths">Allowlist of permitted base paths forwarded to the per-image save step.</param>
    /// <param name="recovery">
    ///     Where this host's publish records live and the key they are signed with (R18-ARCH01).
    /// </param>
    /// <returns>An <see cref="ExtractionStats" /> with extraction counts.</returns>
    private static ExtractionStats ExtractAllImages(Presentation presentation, ExtractionParameters p,
        IReadOnlyList<string> allowedBasePaths, RecoveryContext recovery)
    {
        var count = 0;
        var skippedCount = 0;
        var exportedHashes = new HashSet<string>();

        // Only a picture frame that carries an image becomes a file. Counting every shape refused
        // a deck of a thousand text boxes and no pictures at all (R3-C05). Counting stops one past
        // the cap: beyond that the answer is the same refusal, and a very large deck need not be
        // walked in full to reach it.
        //
        // When duplicates are being skipped, the count is of the files that would actually be
        // written: a deck of five thousand frames sharing one image produces one file, and
        // refusing it for the frame count refused a request that was never going to be large
        // (R4-R06).
        RenderBudget.EnsureOutputCount(
            CountExtractableImages(presentation, p.SkipDuplicates), "image files");

        // A file count says nothing about what lands on disk, so the running total is checked as
        // it grows and each file is written through what is left of it (R4-R07). The batch is
        // published once, at the end: an extraction refused part-way used to leave whichever
        // images it had already produced on top of the caller's files (R5-R01).
        using var batch = new BoundedFileBatch(RenderBudget.MaxOutputBytes, "extracted images",
            recovery, allowedBasePaths);

        var slideNum = 0;
        foreach (var slide in presentation.Slides)
        {
            slideNum++;
            foreach (var shape in slide.Shapes)
                if (shape is PictureFrame { PictureFormat.Picture.Image: not null } pic)
                {
                    var extracted = TryExtractImage(pic, slideNum, ref count, p, exportedHashes,
                        allowedBasePaths, batch);
                    if (!extracted) skippedCount++;
                    RenderBudget.EnsureOutputCount(count, "image files");
                }
        }

        batch.Publish();

        return new ExtractionStats(count, skippedCount);
    }

    /// <summary>
    ///     Counts the picture frames that would produce a file, stopping one past the cap.
    /// </summary>
    /// <param name="presentation">The presentation to inspect.</param>
    /// <param name="skipDuplicates">
    ///     Whether frames sharing an image will produce one file between them, in which case they
    ///     count once.
    /// </param>
    /// <returns>
    ///     The number of files the extraction would write, or
    ///     <see cref="RenderBudget.MaxOutputFiles" /> + 1 once the cap is known to be exceeded.
    /// </returns>
    private static int CountExtractableImages(Presentation presentation, bool skipDuplicates)
    {
        var extractable = 0;
        var seen = skipDuplicates ? new HashSet<string>() : null;

        foreach (var slide in presentation.Slides)
        foreach (var shape in slide.Shapes)
        {
            if (shape is not PictureFrame { PictureFormat.Picture.Image: not null } picture) continue;

            // Frames sharing an image produce one file between them when duplicates are skipped,
            // so counting frames refused requests that would have written a handful of files
            // (R4-R06). Hashing here costs the same work the extraction would do anyway.
            if (seen != null && !seen.Add(PptImageHelper.ComputeImageHash(
                    picture.PictureFormat.Picture.Image.BinaryData)))
                continue;

            if (++extractable > RenderBudget.MaxOutputFiles) return extractable;
        }

        return extractable;
    }

    /// <summary>
    ///     Tries to extract an image from a picture frame.
    /// </summary>
    /// <param name="pic">The picture frame.</param>
    /// <param name="slideNum">The slide number.</param>
    /// <param name="count">The current extraction count.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <param name="exportedHashes">The set of exported image hashes.</param>
    /// <param name="allowedBasePaths">Allowlist of permitted base paths used to validate the resolved output path.</param>
    /// <param name="batch">The request's staged output, published only once all of it succeeds.</param>
    /// <returns>True if the image was extracted, false if skipped.</returns>
    /// <exception cref="ArgumentException">Thrown when the request has written all it may.</exception>
    private static bool TryExtractImage(PictureFrame pic, int slideNum, ref int count,
        ExtractionParameters p, HashSet<string> exportedHashes, IReadOnlyList<string> allowedBasePaths,
        BoundedFileBatch batch)
    {
        var image = pic.PictureFormat.Picture.Image;

        if (p.SkipDuplicates)
        {
            var hash = PptImageHelper.ComputeImageHash(image.BinaryData);
            if (!exportedHashes.Add(hash))
                return false;
        }

        var fileName = Path.Combine(p.OutputDir, $"slide{slideNum}_img{++count}.{p.Extension}");
        // H25: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        fileName = SecurityHelper.ResolveAndEnsureWithinAllowlist(fileName, allowedBasePaths, nameof(fileName));

        // Written through what is left of the budget: saving straight to the destination meant
        // this path had a file count and no byte limit at all, so one deck of very large images
        // was unbounded (R4-R07).
        batch.Stage(fileName,
            stream => image.SystemImage.Save(stream, p.IsJpeg ? ImageFormat.Jpeg : ImageFormat.Png));
        return true;
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    /// <param name="count">The number of extracted images.</param>
    /// <param name="skippedCount">The number of skipped duplicates.</param>
    /// <param name="p">The extraction parameters.</param>
    /// <returns>The result message.</returns>
    private static string BuildResultMessage(int count, int skippedCount, ExtractionParameters p)
    {
        var result = $"Extracted {count} images. Output: {Path.GetFullPath(p.OutputDir)}";
        if (p.SkipDuplicates && skippedCount > 0)
            result += $" (skipped {skippedCount} duplicates)";
        return result;
    }

    /// <summary>
    ///     Record for holding extraction parameters.
    /// </summary>
    /// <param name="OutputDir">The output directory.</param>
    /// <param name="IsJpeg">Whether the output format is JPEG (false = PNG).</param>
    /// <param name="Extension">The file extension.</param>
    /// <param name="SkipDuplicates">Whether to skip duplicate images.</param>
    private sealed record ExtractionParameters(
        string OutputDir,
        bool IsJpeg,
        string Extension,
        bool SkipDuplicates);
}
