using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using ModelContextProtocol;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for merging multiple Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MergeWordDocumentsHandler : OperationHandlerBase<Document>
{
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
        if (string.IsNullOrEmpty(p.OutputPath))
            throw new ArgumentException("outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(p.OutputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        foreach (var inputPath in p.InputPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        var importFormatMode = p.ImportFormatModeStr switch
        {
            "UseDestinationStyles" => ImportFormatMode.UseDestinationStyles,
            "KeepDifferentStyles" => ImportFormatMode.KeepDifferentStyles,
            _ => ImportFormatMode.KeepSourceFormatting
        };

        var mergedDoc = new Document(p.InputPaths[0]);
        var totalFiles = p.InputPaths.Length;

        var initialProgress = 100 / totalFiles;
        context.Progress?.Report(new ProgressNotificationValue
            { Progress = initialProgress, Total = 100, Message = $"Loaded document 1 of {totalFiles}" });

        for (var i = 1; i < p.InputPaths.Length; i++)
        {
            var doc = new Document(p.InputPaths[i]);
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

        mergedDoc.Save(p.OutputPath);
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
            parameters.GetOptional<string[]?>("inputPaths"),
            parameters.GetOptional<string?>("outputPath"),
            parameters.GetOptional("importFormatMode", "KeepSourceFormatting"),
            parameters.GetOptional("unlinkHeadersFooters", false));
    }

    private sealed record MergeParameters(
        string[]? InputPaths,
        string? OutputPath,
        string ImportFormatModeStr,
        bool UnlinkHeadersFooters);
}
