using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for updating fields in Word documents.
/// </summary>
public class UpdateFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "update_field";

    /// <summary>
    ///     Updates one or all fields in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: fieldIndex (update specific field)
    ///     Optional: updateAll (update all fields)
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var updateAll = parameters.GetOptional<bool?>("updateAll");

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.HasValue && updateAll != true)
        {
            if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex.Value];
            if (field.IsLocked)
                return $"Warning: Field #{fieldIndex.Value} is locked and cannot be updated.";

            var oldResult = field.Result ?? "";
            field.Update();
            var newResult = field.Result ?? "";

            MarkModified(context);

            return $"Field #{fieldIndex.Value} updated\nOld result: {oldResult}\nNew result: {newResult}";
        }

        var lockedCount = fields.Count(f => f.IsLocked);
        document.UpdateFields();
        MarkModified(context);

        var result = $"Updated {fields.Count - lockedCount} field(s)";
        if (lockedCount > 0)
            result += $"\nSkipped {lockedCount} locked field(s)";
        return result;
    }
}
