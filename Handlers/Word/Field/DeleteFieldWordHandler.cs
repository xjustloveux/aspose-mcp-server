using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for deleting fields from Word documents.
/// </summary>
public class DeleteFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_field";

    /// <summary>
    ///     Deletes a field from the document, optionally keeping its result text.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldIndex
    ///     Optional: keepResult (default: false)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var keepResult = parameters.GetOptional("keepResult", false);

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for delete_field operation");

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];
        var fieldType = field.Type.ToString();
        var fieldCodeStr = field.GetFieldCode();

        if (keepResult)
            field.Unlink();
        else
            field.Remove();

        MarkModified(context);

        var remainingFields = document.Range.Fields.Count;
        var result = $"Field #{fieldIndex.Value} deleted successfully\n";
        result += $"Type: {fieldType}\nCode: {fieldCodeStr}\n";
        result += $"Keep result text: {(keepResult ? "Yes" : "No")}\n";
        result += $"Remaining fields: {remainingFields}";
        return result;
    }
}
