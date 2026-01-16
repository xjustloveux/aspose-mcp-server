using Aspose.Words;
using Aspose.Words.Comparing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Handler for comparing two Word documents.
/// </summary>
public class CompareDocumentsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "compare";

    /// <summary>
    ///     Compares two documents and creates a comparison document showing differences as revisions.
    /// </summary>
    /// <param name="context">The document context (not used, comparison creates new document).</param>
    /// <param name="parameters">
    ///     Required: originalPath, revisedPath, outputPath
    ///     Optional: authorName (default: "Comparison"), ignoreFormatting (default: false),
    ///     ignoreComments (default: false)
    /// </param>
    /// <returns>Success message with comparison result and number of differences found.</returns>
    /// <exception cref="ArgumentException">Thrown when required paths are not provided.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractCompareDocumentsParameters(parameters);

        var originalDoc = new Document(p.OriginalPath);
        var revisedDoc = new Document(p.RevisedPath);

        var compareOptions = new CompareOptions
        {
            IgnoreFormatting = p.IgnoreFormatting,
            IgnoreComments = p.IgnoreComments
        };

        originalDoc.Compare(revisedDoc, p.AuthorName, DateTime.Now, compareOptions);
        var revisionCount = originalDoc.Revisions.Count;
        originalDoc.Save(p.OutputPath);

        return Success($"Comparison completed: {revisionCount} difference(s) found\nOutput: {p.OutputPath}");
    }

    private static CompareDocumentsParameters ExtractCompareDocumentsParameters(OperationParameters parameters)
    {
        var outputPath = parameters.GetRequired<string>("outputPath");
        var originalPath = parameters.GetRequired<string>("originalPath");
        var revisedPath = parameters.GetRequired<string>("revisedPath");

        SecurityHelper.ValidateFilePath(originalPath, "originalPath", true);
        SecurityHelper.ValidateFilePath(revisedPath, "revisedPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return new CompareDocumentsParameters(
            originalPath,
            revisedPath,
            outputPath,
            parameters.GetOptional("authorName", "Comparison"),
            parameters.GetOptional("ignoreFormatting", false),
            parameters.GetOptional("ignoreComments", false));
    }

    private record CompareDocumentsParameters(
        string OriginalPath,
        string RevisedPath,
        string OutputPath,
        string AuthorName,
        bool IgnoreFormatting,
        bool IgnoreComments);
}
