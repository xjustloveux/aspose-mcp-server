using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for merging multiple PowerPoint presentations.
/// </summary>
public class MergePresentationsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges multiple PowerPoint presentations into one.
    /// </summary>
    /// <param name="context">The presentation context (not used for merge).</param>
    /// <param name="parameters">
    ///     Required: inputPaths, path or outputPath
    ///     Optional: keepSourceFormatting
    /// </param>
    /// <returns>Success message with output path and slide count.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var path = parameters.GetOptional<string?>("path");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var inputPaths = parameters.GetRequired<string[]>("inputPaths");
        var keepSourceFormatting = parameters.GetOptional("keepSourceFormatting", true);

        var savePath = path ?? outputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(savePath, "outputPath", true);

        var validPaths = inputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        using var masterPresentation = new Presentation(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            var inputPath = validPaths[i];
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath)) continue;

            using var sourcePresentation = new Presentation(inputPath);
            foreach (var slide in sourcePresentation.Slides)
                if (keepSourceFormatting)
                {
                    var sourceMaster = slide.LayoutSlide.MasterSlide;
                    var destMaster = masterPresentation.Masters.AddClone(sourceMaster);
                    masterPresentation.Slides.AddClone(slide, destMaster, true);
                }
                else
                {
                    masterPresentation.Slides.AddClone(slide, masterPresentation.Masters[0], true);
                }
        }

        masterPresentation.Save(savePath, SaveFormat.Pptx);

        return Success(
            $"Merged {validPaths.Count} presentations (Total slides: {masterPresentation.Slides.Count}). Output: {savePath}");
    }
}
