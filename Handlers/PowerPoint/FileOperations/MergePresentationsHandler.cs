using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Handler for merging multiple PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MergePresentationsHandler : OperationHandlerBase<Presentation>
{
    private const string InputPathsParamName = "inputPaths";

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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        // Held for the whole operation: this handler builds or reads presentations of
        // its own, outside any session, so nothing else stands between it and a library
        // that fails when two threads are inside it (SlidesGate).
        using var slidesGate = SlidesGate.Enter();

        var p = ExtractMergeParameters(parameters);

        var savePath = p.Path ?? p.OutputPath;
        if (string.IsNullOrEmpty(savePath))
            throw new ArgumentException("path or outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(savePath, "outputPath", true);
        savePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(savePath,
            context.ServerConfig?.AllowedBasePaths ?? [], "outputPath");

        // One request should not be able to pull in an unbounded number of documents; PDF
        // merge has capped this since RB-29 and the other three families had not.
        SecurityHelper.ValidateArraySize(p.InputPaths, "inputPaths");

        var validPaths = p.InputPaths.Where(path => !string.IsNullOrEmpty(path)).ToList();

        if (!string.IsNullOrEmpty(p.InputPath))
            validPaths.Insert(0, p.InputPath);

        if (validPaths.Count == 0)
            throw new ArgumentException("No valid input paths provided");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, InputPathsParamName, true);

        // Inputs are read sinks: resolve symlinks and enforce the allowlist immediately before each open.
        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];
        var resolvedMasterPath =
            SecurityHelper.ResolveAndEnsureWithinAllowlist(validPaths[0], allowedBasePaths, InputPathsParamName);
        using var masterPresentation = new Presentation(resolvedMasterPath);

        for (var i = 1; i < validPaths.Count; i++)
        {
            var inputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(validPaths[i], allowedBasePaths,
                InputPathsParamName);
            if (!File.Exists(inputPath)) continue;

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

        // H20: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        savePath = SecurityHelper.ResolveAndEnsureWithinAllowlist(savePath,
            context.ServerConfig?.AllowedBasePaths ?? [], nameof(savePath));
        masterPresentation.Save(savePath, PptSaveFormatResolver.Resolve(savePath));

        return new SuccessResult
        {
            Message =
                $"Merged {validPaths.Count} presentations (Total slides: {masterPresentation.Slides.Count}). Output: {savePath}"
        };
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
            parameters.GetOptional<string?>("inputPath"),
            parameters.GetRequired<string[]>(InputPathsParamName),
            parameters.GetOptional("keepSourceFormatting", true));
    }

    /// <summary>
    ///     Record for holding merge presentations parameters.
    /// </summary>
    /// <param name="Path">The output file path.</param>
    /// <param name="OutputPath">Alternative output file path.</param>
    /// <param name="InputPath">The base presentation file path.</param>
    /// <param name="InputPaths">The array of input file paths to merge.</param>
    /// <param name="KeepSourceFormatting">Whether to keep source formatting.</param>
    private sealed record MergeParameters(
        string? Path,
        string? OutputPath,
        string? InputPath,
        string[] InputPaths,
        bool KeepSourceFormatting);
}
