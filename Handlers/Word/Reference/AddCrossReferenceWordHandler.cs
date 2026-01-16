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
        var p = ExtractAddCrossReferenceParameters(parameters);

        var validTypes = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" };
        if (!validTypes.Contains(p.ReferenceType, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException(
                $"Invalid referenceType: {p.ReferenceType}. Valid types are: {string.Join(", ", validTypes)}");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(p.ReferenceText))
            builder.Write(p.ReferenceText);

        var fieldCode = p.InsertAsHyperlink ? $"REF {p.TargetName} \\h" : $"REF {p.TargetName}";

        builder.InsertField(fieldCode);
        if (p.IncludeAboveBelow)
            builder.Write(" (above)");

        MarkModified(context);

        return Success($"Cross-reference added (Type: {p.ReferenceType})");
    }

    private static AddCrossReferenceParameters ExtractAddCrossReferenceParameters(OperationParameters parameters)
    {
        var referenceType = parameters.GetOptional<string?>("referenceType");
        var targetName = parameters.GetOptional<string?>("targetName");

        if (string.IsNullOrEmpty(referenceType))
            throw new ArgumentException("referenceType is required for add_cross_reference operation");
        if (string.IsNullOrEmpty(targetName))
            throw new ArgumentException("targetName is required for add_cross_reference operation");

        return new AddCrossReferenceParameters(
            referenceType,
            targetName,
            parameters.GetOptional<string?>("referenceText"),
            parameters.GetOptional("insertAsHyperlink", true),
            parameters.GetOptional("includeAboveBelow", false));
    }

    private record AddCrossReferenceParameters(
        string ReferenceType,
        string TargetName,
        string? ReferenceText,
        bool InsertAsHyperlink,
        bool IncludeAboveBelow);
}
