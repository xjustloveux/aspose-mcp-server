using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for splitting a PowerPoint presentation into multiple files.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SplitPresentationHandler : OperationHandlerBase<Presentation>
{
    /// <summary>
    ///     Maximum allowed total work units (slideCount × outputFileCount) per split call.
    ///     This is orthogonal to the 1000-file output cap: a caller requesting 999 files
    ///     from a 200-slide deck would produce 199,800 work units and be rejected here even
    ///     though the file count is under 1000.  The bound was chosen so that a typical
    ///     1000-slide deck split into 100 files (100,000 units) is at the limit, while
    ///     adversarial combinations (e.g., 10,000 slides × 999 files) are blocked.
    /// </summary>
    private const int MaxTotalWorkUnits = 100_000;

    /// <inheritdoc />
    public override string Operation => "split";

    /// <summary>
    ///     Splits a PowerPoint presentation into multiple files.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: outputDirectory
    ///     Optional: inputPath, path, sessionId, slidesPerFile (1..1000, default 1), startSlideIndex, endSlideIndex,
    ///     outputFileNamePattern
    /// </param>
    /// <returns>Success message with output directory and file count.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when <c>slidesPerFile</c> is outside [1,1000]; when neither
    ///     <c>inputPath</c>/<c>path</c> nor <c>sessionId</c> is supplied; when the
    ///     slide range (<c>startSlideIndex</c>/<c>endSlideIndex</c>) is invalid;
    ///     when the computed output file count would exceed 1000; when the product
    ///     of slide range and output file count exceeds <see cref="MaxTotalWorkUnits" />;
    ///     when <c>outputFileNamePattern</c> exceeds the per-filename length cap or is
    ///     missing the <c>{index}</c> placeholder; when, after sanitation and
    ///     placeholder substitution, the per-slide output path resolves outside
    ///     <c>outputDirectory</c>; or when path validation fails.
    /// </exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when a <c>sessionId</c> is supplied but session management is disabled
    ///     on the context.
    /// </exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSplitParameters(parameters);

        if (p.SlidesPerFile < 1 || p.SlidesPerFile > 1000)
            throw new ArgumentException("slidesPerFile must be between 1 and 1000");

        // Reject the raw pattern before sanitation — the sanitizer silently truncates
        // to 255 chars, which would strip a late-positioned {index} and regress into
        // the silent-overwrite symptom this validation exists to close.
        if (p.OutputFileNamePattern.Length > 255)
            throw new ArgumentException("outputFileNamePattern exceeds maximum length");
        if (!p.OutputFileNamePattern.Contains("{index}", StringComparison.Ordinal))
            throw new ArgumentException("outputFileNamePattern must contain the {index} placeholder");

        var sourcePath = p.InputPath ?? p.Path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required for split operation");

        if (!string.IsNullOrEmpty(sourcePath))
            SecurityHelper.ValidateFilePath(sourcePath, "inputPath", true);
        SecurityHelper.ValidateFilePath(p.OutputDirectory, "outputDirectory", true);

        if (!Directory.Exists(p.OutputDirectory))
            Directory.CreateDirectory(p.OutputDirectory);

        Presentation presentation;
        Presentation? ownedPresentation = null;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            presentation = context.SessionManager.GetDocument<Presentation>(p.SessionId, identity);
        }
        else
        {
            ownedPresentation = new Presentation(sourcePath);
            presentation = ownedPresentation;
        }

        var totalSlides = presentation.Slides.Count;

        try
        {
            var start = p.StartSlideIndex ?? 0;
            var end = p.EndSlideIndex ?? totalSlides - 1;

            if (start < 0 || start >= totalSlides || end < 0 || end >= totalSlides || start > end)
                throw new ArgumentException($"Invalid slide range: start={start}, end={end}, total={totalSlides}");

            var outputFileCount = (end - start + p.SlidesPerFile) / p.SlidesPerFile;
            if (outputFileCount > 1000)
                throw new ArgumentException(
                    "Split would produce too many files (max 1000). " +
                    "Increase slidesPerFile or narrow startSlideIndex / endSlideIndex.");

            // Orthogonal DoS guard: even when outputFileCount ≤ 1000, a large slide range
            // combined with many output files causes excessive I/O.  Each slide copied to each
            // output file counts as one work unit.
            var slideRange = end - start + 1;
            var totalWorkUnits = (long)slideRange * outputFileCount;
            if (totalWorkUnits > MaxTotalWorkUnits)
                throw new ArgumentException(
                    "Split operation exceeds the maximum allowed work units (100000). Reduce the slide range or increase slidesPerFile.");

            // Sanitize the pattern once; it is invariant per call and only {index} varies.
            var safePattern = SecurityHelper.SanitizeFileNamePattern(p.OutputFileNamePattern);
            // Normalize for prefix comparison (trailing sep avoids "/out2" matching "/out").
            var normalizedOutputDir =
                Path.GetFullPath(p.OutputDirectory)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            var fileCount = 0;
            for (var i = start; i <= end; i += p.SlidesPerFile)
            {
                using var newPresentation = new Presentation();
                newPresentation.Slides.RemoveAt(0);

                for (var j = 0; j < p.SlidesPerFile && i + j <= end; j++)
                {
                    var sourceSlide = presentation.Slides[i + j];
                    var sourceMaster = sourceSlide.LayoutSlide.MasterSlide;
                    var destMaster = newPresentation.Masters.AddClone(sourceMaster);
                    newPresentation.Slides.AddClone(sourceSlide, destMaster, true);
                }

                var outputFileName = safePattern.Replace("{index}", fileCount.ToString());
                var outPath = Path.Combine(p.OutputDirectory, outputFileName);
                if (!Path.GetFullPath(outPath)
                        .StartsWith(normalizedOutputDir, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        "outputFileNamePattern resolves to a path outside outputDirectory");
                // H19: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
                outPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(outPath,
                    context.ServerConfig?.AllowedBasePaths ?? [], nameof(outPath));
                newPresentation.Save(outPath, SaveFormat.Pptx);
                fileCount++;
            }

            return new SuccessResult
                { Message = $"Split presentation into {fileCount} file(s). Output: {p.OutputDirectory}" };
        }
        finally
        {
            ownedPresentation?.Dispose();
        }
    }

    /// <summary>
    ///     Extracts split parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted split parameters.</returns>
    private static SplitParameters ExtractSplitParameters(OperationParameters parameters)
    {
        return new SplitParameters(
            parameters.GetOptional<string?>("inputPath"),
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("sessionId"),
            parameters.GetRequired<string>("outputDirectory"),
            parameters.GetOptional("slidesPerFile", 1),
            parameters.GetOptional<int?>("startSlideIndex"),
            parameters.GetOptional<int?>("endSlideIndex"),
            parameters.GetOptional("outputFileNamePattern", "slide_{index}.pptx"));
    }

    /// <summary>
    ///     Record for holding split presentation parameters.
    /// </summary>
    /// <param name="InputPath">The input file path.</param>
    /// <param name="Path">Alternative input file path.</param>
    /// <param name="SessionId">The session ID.</param>
    /// <param name="OutputDirectory">The output directory.</param>
    /// <param name="SlidesPerFile">Number of slides per output file.</param>
    /// <param name="StartSlideIndex">The starting slide index.</param>
    /// <param name="EndSlideIndex">The ending slide index.</param>
    /// <param name="OutputFileNamePattern">The output file name pattern.</param>
    private sealed record SplitParameters(
        string? InputPath,
        string? Path,
        string? SessionId,
        string OutputDirectory,
        int SlidesPerFile,
        int? StartSlideIndex,
        int? EndSlideIndex,
        string OutputFileNamePattern);
}
