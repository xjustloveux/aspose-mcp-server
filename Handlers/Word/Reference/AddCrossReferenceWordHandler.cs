using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for adding cross-references to Word documents.
/// </summary>
public class AddCrossReferenceWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_cross_reference";

    /// <summary>
    ///     Adds a cross-reference (REF field) to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: referenceType, targetName
    ///     Optional: referenceText, insertAsHyperlink (default: true), includeAboveBelow (default: false)
    /// </param>
    /// <returns>Success message indicating cross-reference was added.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var referenceType = parameters.GetOptional<string?>("referenceType");
        var referenceText = parameters.GetOptional<string?>("referenceText");
        var targetName = parameters.GetOptional<string?>("targetName");
        var insertAsHyperlink = parameters.GetOptional("insertAsHyperlink", true);
        var includeAboveBelow = parameters.GetOptional("includeAboveBelow", false);

        if (string.IsNullOrEmpty(referenceType))
            throw new ArgumentException("referenceType is required for add_cross_reference operation");
        if (string.IsNullOrEmpty(targetName))
            throw new ArgumentException("targetName is required for add_cross_reference operation");

        var validTypes = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" };
        if (!validTypes.Contains(referenceType, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Invalid referenceType: {referenceType}. Valid types are: {string.Join(", ", validTypes)}");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(referenceText))
            builder.Write(referenceText);

        var fieldCode = insertAsHyperlink ? $"REF {targetName} \\h" : $"REF {targetName}";

        builder.InsertField(fieldCode);
        if (includeAboveBelow)
            builder.Write(" (above)");

        MarkModified(context);

        return Success($"Cross-reference added (Type: {referenceType})");
    }
}
