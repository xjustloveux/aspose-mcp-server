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
        var p = ExtractMergeParameters(parameters);

        var savePath = p.Path ?? p.OutputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(savePath, "outputPath", true);

        var validPaths = p.InputPaths.Where(path => !string.IsNullOrEmpty(path)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        using var masterPresentation = new Presentation(validPaths[0]);

        for (var i = 1; i < validPaths.Count; i++)
        {
            var inputPath = validPaths[i];
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath)) continue;

            using var sourcePresentation = new Presentation(inputPath);
            foreach (var slide in sourcePresentation.Slides)
                if (p.KeepSourceFormatting)
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

    /// <summary>
    ///     Extracts merge parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted merge parameters.</returns>
    private static MergeParameters ExtractMergeParameters(OperationParameters parameters)
    {
        return new MergeParameters(
            parameters.GetOptional<string?>("path"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetRequired<string[]>("inputPaths"),
            parameters.GetOptional("keepSourceFormatting", true));
    }

    /// <summary>
    ///     Record for holding merge presentations parameters.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path.</param>
    /// <param name="InputPaths">The array of input file paths.</param>
    /// <param name="KeepSourceFormatting">Whether to keep source formatting.</param>
    private sealed record MergeParameters(
        string? Path,
        string? OutputPath,
        string[] InputPaths,
        bool KeepSourceFormatting);
}
