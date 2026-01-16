using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for splitting a PowerPoint presentation into multiple files.
/// </summary>
public class SplitPresentationHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "split";

    /// <summary>
    ///     Splits a PowerPoint presentation into multiple files.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: outputDirectory
    ///     Optional: inputPath, path, sessionId, slidesPerFile, startSlideIndex, endSlideIndex, outputFileNamePattern
    /// </param>
    /// <returns>Success message with output directory and file count.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSplitParameters(parameters);

        var sourcePath = p.InputPath ?? p.Path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(p.SessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required for split operation");

        if (!Directory.Exists(p.OutputDirectory))
            Directory.CreateDirectory(p.OutputDirectory);

        Presentation presentation;

        if (!string.IsNullOrEmpty(p.SessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            presentation = context.SessionManager.GetDocument<Presentation>(p.SessionId, identity);
        }
        else
        {
            presentation = new Presentation(sourcePath);
        }

        var totalSlides = presentation.Slides.Count;

        var start = p.StartSlideIndex ?? 0;
        var end = p.EndSlideIndex ?? totalSlides - 1;

        if (start < 0 || start >= totalSlides || end < 0 || end >= totalSlides || start > end)
            throw new ArgumentException($"Invalid slide range: start={start}, end={end}, total={totalSlides}");

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

            var outputFileName = p.OutputFileNamePattern.Replace("{index}", fileCount.ToString());
            outputFileName = SecurityHelper.SanitizeFileName(outputFileName);
            var outPath = Path.Combine(p.OutputDirectory, outputFileName);
            newPresentation.Save(outPath, SaveFormat.Pptx);
            fileCount++;
        }

        return Success($"Split presentation into {fileCount} file(s). Output: {p.OutputDirectory}");
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
    private record SplitParameters(
        string? InputPath,
        string? Path,
        string? SessionId,
        string OutputDirectory,
        int SlidesPerFile,
        int? StartSlideIndex,
        int? EndSlideIndex,
        string OutputFileNamePattern);
}
