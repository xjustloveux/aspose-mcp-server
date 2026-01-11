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
        var inputPath = parameters.GetOptional<string?>("inputPath");
        var path = parameters.GetOptional<string?>("path");
        var sessionId = parameters.GetOptional<string?>("sessionId");
        var outputDirectory = parameters.GetRequired<string>("outputDirectory");
        var slidesPerFile = parameters.GetOptional("slidesPerFile", 1);
        var startSlideIndex = parameters.GetOptional<int?>("startSlideIndex");
        var endSlideIndex = parameters.GetOptional<int?>("endSlideIndex");
        var outputFileNamePattern = parameters.GetOptional("outputFileNamePattern", "slide_{index}.pptx");

        var sourcePath = inputPath ?? path;
        if (string.IsNullOrEmpty(sourcePath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath, path, or sessionId is required for split operation");

        if (!Directory.Exists(outputDirectory))
            Directory.CreateDirectory(outputDirectory);

        Presentation presentation;

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (context.SessionManager == null)
                throw new InvalidOperationException("Session management is not enabled");

            var identity = context.IdentityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
            presentation = context.SessionManager.GetDocument<Presentation>(sessionId, identity);
        }
        else
        {
            presentation = new Presentation(sourcePath);
        }

        var totalSlides = presentation.Slides.Count;

        var start = startSlideIndex ?? 0;
        var end = endSlideIndex ?? totalSlides - 1;

        if (start < 0 || start >= totalSlides || end < 0 || end >= totalSlides || start > end)
            throw new ArgumentException($"Invalid slide range: start={start}, end={end}, total={totalSlides}");

        var fileCount = 0;
        for (var i = start; i <= end; i += slidesPerFile)
        {
            using var newPresentation = new Presentation();
            newPresentation.Slides.RemoveAt(0);

            for (var j = 0; j < slidesPerFile && i + j <= end; j++)
            {
                var sourceSlide = presentation.Slides[i + j];
                var sourceMaster = sourceSlide.LayoutSlide.MasterSlide;
                var destMaster = newPresentation.Masters.AddClone(sourceMaster);
                newPresentation.Slides.AddClone(sourceSlide, destMaster, true);
            }

            var outputFileName = outputFileNamePattern.Replace("{index}", fileCount.ToString());
            outputFileName = SecurityHelper.SanitizeFileName(outputFileName);
            var outPath = Path.Combine(outputDirectory, outputFileName);
            newPresentation.Save(outPath, SaveFormat.Pptx);
            fileCount++;
        }

        return Success($"Split presentation into {fileCount} file(s). Output: {outputDirectory}");
    }
}
