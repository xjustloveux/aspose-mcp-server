using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for deleting fields from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteFieldParameters(parameters);

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (p.FieldIndex < 0 || p.FieldIndex >= fields.Count)
            throw new ArgumentException(
                $"Field index {p.FieldIndex} is out of range (document has {fields.Count} fields)");

        var field = fields[p.FieldIndex];
        var fieldType = field.Type.ToString();
        var fieldCodeStr = field.GetFieldCode();

        if (p.KeepResult)
            field.Unlink();
        else
            field.Remove();

        MarkModified(context);

        var remainingFields = document.Range.Fields.Count;
        var message = $"Field #{p.FieldIndex} deleted successfully\n";
        message += $"Type: {fieldType}\nCode: {fieldCodeStr}\n";
        message += $"Keep result text: {(p.KeepResult ? "Yes" : "No")}\n";
        message += $"Remaining fields: {remainingFields}";
        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts and validates parameters for the delete field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided.</exception>
    private static DeleteFieldParameters ExtractDeleteFieldParameters(OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var keepResult = parameters.GetOptional("keepResult", false);

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for delete_field operation");

        return new DeleteFieldParameters(fieldIndex.Value, keepResult);
    }

    /// <summary>
    ///     Parameters for the delete field operation.
    /// </summary>
    /// <param name="FieldIndex">The index of the field to delete.</param>
    /// <param name="KeepResult">Whether to keep the field's result text after deletion.</param>
    private sealed record DeleteFieldParameters(int FieldIndex, bool KeepResult);
}
