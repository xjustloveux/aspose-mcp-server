using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Handler for merging multiple Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var inputPaths = parameters.GetOptional<string[]?>("inputPaths");
        var outputPath = parameters.GetOptional<string?>("outputPath");
        var importFormatModeStr = parameters.GetOptional("importFormatMode", "KeepSourceFormatting");
        var unlinkHeadersFooters = parameters.GetOptional("unlinkHeadersFooters", false);

        if (inputPaths == null || inputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required for merge operation");

        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir))
            Directory.CreateDirectory(outputDir);

        foreach (var inputPath in inputPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);

        var importFormatMode = importFormatModeStr switch
        {
            "UseDestinationStyles" => ImportFormatMode.UseDestinationStyles,
            "KeepDifferentStyles" => ImportFormatMode.KeepDifferentStyles,
            _ => ImportFormatMode.KeepSourceFormatting
        };

        var mergedDoc = new Document(inputPaths[0]);

        for (var i = 1; i < inputPaths.Length; i++)
        {
            var doc = new Document(inputPaths[i]);
            mergedDoc.AppendDocument(doc, importFormatMode);
        }

        if (unlinkHeadersFooters)
            foreach (var section in mergedDoc.Sections.Cast<Section>())
                section.HeadersFooters.LinkToPrevious(false);

        mergedDoc.Save(outputPath);
        return $"Merged {inputPaths.Length} documents into: {outputPath} (format mode: {importFormatModeStr})";
    }
}
