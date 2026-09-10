using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for updating table of contents in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class UpdateTableOfContentsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "update_toc";

    /// <summary>
    ///     Updates the table of contents fields in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tocIndex (0-based index of specific TOC to update)
    /// </param>
    /// <returns>Success message or info message if no TOC found.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractUpdateTableOfContentsParameters(parameters);

        var doc = context.Document;
        var tocFields = doc.Range.Fields
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        if (tocFields.Count == 0)
        {
            var allFields = doc.Range.Fields.ToList();
            var fieldTypes = allFields.Select(f => f.Type.ToString()).Distinct().ToList();
            var message = "No table of contents fields found in document.";
            if (allFields.Count > 0)
                message += $" Found {allFields.Count} field(s) of other types: {string.Join(", ", fieldTypes)}.";
            message += " Use 'add_toc' operation to add a table of contents first.";
            return new SuccessResult { Message = message };
        }

        // This handler updates the tables of contents and then every other field in the
        // document. Both passes are checked before either runs: refusing at the second one left
        // the tables of contents already rebuilt in a document the session keeps (R7-W01).
        if (p.TocIndex.HasValue && (p.TocIndex.Value < 0 || p.TocIndex.Value >= tocFields.Count))
            throw new ArgumentException($"tocIndex must be between 0 and {tocFields.Count - 1}");

        WordFieldPolicy.RefuseNestedDisallowedFields(doc, null);

        if (p.TocIndex.HasValue)
            WordFieldPolicy.UpdateField(tocFields[p.TocIndex.Value]);
        else
            foreach (var tocField in tocFields)
                WordFieldPolicy.UpdateField(tocField);

        WordFieldPolicy.UpdateAllowedFields(doc);

        MarkModified(context);

        var updatedCount = p.TocIndex.HasValue ? 1 : tocFields.Count;
        return new SuccessResult { Message = $"Updated {updatedCount} table of contents field(s)" };
    }

    private static UpdateTableOfContentsParameters ExtractUpdateTableOfContentsParameters(
        OperationParameters parameters)
    {
        return new UpdateTableOfContentsParameters(
            parameters.GetOptional<int?>("tocIndex"));
    }

    private sealed record UpdateTableOfContentsParameters(int? TocIndex);
}
