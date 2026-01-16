using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Handler for updating table of contents in Word documents.
/// </summary>
public class UpdateTableOfContentsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "update_table_of_contents";

    /// <summary>
    ///     Updates the table of contents fields in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tocIndex (0-based index of specific TOC to update)
    /// </param>
    /// <returns>Success message or info message if no TOC found.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
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
            message += " Use 'add_table_of_contents' operation to add a table of contents first.";
            return message;
        }

        if (p.TocIndex.HasValue)
        {
            if (p.TocIndex.Value < 0 || p.TocIndex.Value >= tocFields.Count)
                throw new ArgumentException($"tocIndex must be between 0 and {tocFields.Count - 1}");
            tocFields[p.TocIndex.Value].Update();
        }
        else
        {
            foreach (var tocField in tocFields)
                tocField.Update();
        }

        doc.UpdateFields();

        MarkModified(context);

        var updatedCount = p.TocIndex.HasValue ? 1 : tocFields.Count;
        return Success($"Updated {updatedCount} table of contents field(s)");
    }

    private static UpdateTableOfContentsParameters ExtractUpdateTableOfContentsParameters(
        OperationParameters parameters)
    {
        return new UpdateTableOfContentsParameters(
            parameters.GetOptional<int?>("tocIndex"));
    }

    private record UpdateTableOfContentsParameters(int? TocIndex);
}
