using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using ModelContextProtocol;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for merging multiple Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MergeWordDocumentsHandler : OperationHandlerBase<Document>
{
    private const string InputPathsParameter = "inputPaths";
    private const string OutputPathParameter = "outputPath";

    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges multiple Word documents into one.
    /// </summary>
    /// <param name="context">The operation context.</param>
    /// <param name="parameters">
    ///     Required: inputPaths, outputPath
    ///     Optional: importFormatMode (default: KeepSourceFormatting), unlinkHeadersFooters (default: false)
    /// </param>
    /// <returns>Success message with merge details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMergeParameters(parameters);

        if (p.InputPaths == null || p.InputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");

        // One request should not be able to pull in an unbounded number of documents; PDF
        // merge has capped this since RB-29 and the other three families had not.
        SecurityHelper.ValidateArraySize(p.InputPaths, InputPathsParameter);
        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, OutputPathParameter, true);
        var outputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        // The canonical paths are kept, not thrown away: validating one path and then
        // opening a different, still-mutable one is the check-to-use gap the resolver
        // exists to close (R7-F03).
        var inputPaths = new List<string>(p.InputPaths.Length);
        foreach (var inputPath in p.InputPaths)
        {
            SecurityHelper.ValidateFilePath(inputPath, InputPathsParameter, true);
            inputPaths.Add(SecurityHelper.ResolveAndEnsureWithinAllowlist(inputPath,
                context.ServerConfig?.AllowedBasePaths ?? [], InputPathsParameter));
        }

        var importFormatMode = p.ImportFormatModeStr switch
        {
            "UseDestinationStyles" => ImportFormatMode.UseDestinationStyles,
            "KeepDifferentStyles" => ImportFormatMode.KeepDifferentStyles,
            _ => ImportFormatMode.KeepSourceFormatting
        };

        var allowedBasePaths = context.ServerConfig?.AllowedBasePaths ?? [];
        var mergedDoc = GuardedWordLoader.Load(inputPaths[0], allowedBasePaths);
        var totalFiles = p.InputPaths.Length;

        var initialProgress = 100 / totalFiles;
        context.Progress?.Report(new ProgressNotificationValue
            { Progress = initialProgress, Total = 100, Message = $"Loaded document 1 of {totalFiles}" });

        for (var i = 1; i < p.InputPaths.Length; i++)
        {
            var doc = GuardedWordLoader.Load(inputPaths[i], allowedBasePaths);
            mergedDoc.AppendDocument(doc, importFormatMode);

            var mergeProgress = (i + 1) * 100 / totalFiles;
            context.Progress?.Report(new ProgressNotificationValue
            {
                Progress = mergeProgress,
                Total = 100,
                Message = $"Merged document {i + 1} of {totalFiles}"
            });
        }

        if (p.UnlinkHeadersFooters)
            foreach (var section in mergedDoc.Sections.Cast<Section>())
                section.HeadersFooters.LinkToPrevious(false);

        // H6: resolve symlinks immediately before the sink (bug 20260415-symlink-toctou-sweep).
        var resolvedOutputPath = SecurityHelper.ResolveAndEnsureWithinAllowlist(p.OutputPath,
            context.ServerConfig?.AllowedBasePaths ?? [], OutputPathParameter);
        mergedDoc.Save(resolvedOutputPath);
        context.Progress?.Report(new ProgressNotificationValue
            { Progress = 100, Total = 100, Message = "Merge completed" });
        return new SuccessResult
        {
            Message =
                $"Merged {p.InputPaths.Length} documents into: {p.OutputPath} (format mode: {p.ImportFormatModeStr})"
        };
    }

    private static MergeParameters ExtractMergeParameters(OperationParameters parameters)
    {
        return new MergeParameters(
            parameters.GetOptional<string[]?>(InputPathsParameter),
            parameters.GetOptional<string?>(OutputPathParameter),
            parameters.GetOptional("importFormatMode", "KeepSourceFormatting"),
            parameters.GetOptional("unlinkHeadersFooters", false));
    }

    private sealed record MergeParameters(
        string[]? InputPaths,
        string? OutputPath,
        string ImportFormatModeStr,
        bool UnlinkHeadersFooters);
}
