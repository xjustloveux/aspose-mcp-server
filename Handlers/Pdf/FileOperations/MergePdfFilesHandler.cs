using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for merging multiple PDF documents into a single document.
/// </summary>
public class MergePdfFilesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges multiple PDF documents into a single document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: outputPath, inputPaths (string array)
    /// </param>
    /// <returns>Success message with merge count and output path.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var mergeParams = ExtractMergeParameters(parameters);

        if (mergeParams.InputPaths.Length == 0)
            throw new ArgumentException("inputPaths is required for merge operation");

        SecurityHelper.ValidateArraySize(mergeParams.InputPaths, "inputPaths");

        var validPaths = mergeParams.InputPaths.Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (validPaths.Count == 0)
            throw new ArgumentException("At least one input path is required");

        foreach (var inputPath in validPaths)
            SecurityHelper.ValidateFilePath(inputPath, "inputPaths", true);
        SecurityHelper.ValidateFilePath(mergeParams.OutputPath, "outputPath", true);

        using var mergedDocument = new Document(validPaths[0]);
        for (var i = 1; i < validPaths.Count; i++)
        {
            using var doc = new Document(validPaths[i]);
            mergedDocument.Pages.Add(doc.Pages);
        }

        mergedDocument.Save(mergeParams.OutputPath);

        return Success($"Merged {validPaths.Count} PDF documents. Output: {mergeParams.OutputPath}");
    }

    /// <summary>
    ///     Extracts merge parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted merge parameters.</returns>
    private static MergeParameters ExtractMergeParameters(OperationParameters parameters)
    {
        return new MergeParameters(
            parameters.GetRequired<string>("outputPath"),
            parameters.GetRequired<string[]>("inputPaths")
        );
    }

    /// <summary>
    ///     Record to hold merge parameters.
    /// </summary>
    private record MergeParameters(string OutputPath, string[] InputPaths);
}
