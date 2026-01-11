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
        var outputPath = parameters.GetRequired<string>("outputPath");
        var originalPath = parameters.GetRequired<string>("originalPath");
        var revisedPath = parameters.GetRequired<string>("revisedPath");
        var authorName = parameters.GetOptional("authorName", "Comparison");
        var ignoreFormatting = parameters.GetOptional("ignoreFormatting", false);
        var ignoreComments = parameters.GetOptional("ignoreComments", false);

        SecurityHelper.ValidateFilePath(originalPath, "originalPath", true);
        SecurityHelper.ValidateFilePath(revisedPath, "revisedPath", true);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        var originalDoc = new Document(originalPath);
        var revisedDoc = new Document(revisedPath);

        var compareOptions = new CompareOptions
        {
            IgnoreFormatting = ignoreFormatting,
            IgnoreComments = ignoreComments
        };

        originalDoc.Compare(revisedDoc, authorName, DateTime.Now, compareOptions);
        var revisionCount = originalDoc.Revisions.Count;
        originalDoc.Save(outputPath);

        return Success($"Comparison completed: {revisionCount} difference(s) found\nOutput: {outputPath}");
    }
}
