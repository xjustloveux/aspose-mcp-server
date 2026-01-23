using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for updating fields in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractUpdateFieldParameters(parameters);

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (p.FieldIndex.HasValue && p.UpdateAll != true)
        {
            if (p.FieldIndex.Value < 0 || p.FieldIndex.Value >= fields.Count)
                throw new ArgumentException(
                    $"Field index {p.FieldIndex.Value} is out of range (document has {fields.Count} fields)");

            var field = fields[p.FieldIndex.Value];
            if (field.IsLocked)
                return new SuccessResult
                    { Message = $"Warning: Field #{p.FieldIndex.Value} is locked and cannot be updated." };

            var oldResult = field.Result ?? "";
            field.Update();
            var newResult = field.Result ?? "";

            MarkModified(context);

            return new SuccessResult
                { Message = $"Field #{p.FieldIndex.Value} updated\nOld result: {oldResult}\nNew result: {newResult}" };
        }

        var lockedCount = fields.Count(f => f.IsLocked);
        document.UpdateFields();
        MarkModified(context);

        var message = $"Updated {fields.Count - lockedCount} field(s)";
        if (lockedCount > 0)
            message += $"\nSkipped {lockedCount} locked field(s)";
        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts parameters for the update field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static UpdateFieldParameters ExtractUpdateFieldParameters(OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var updateAll = parameters.GetOptional<bool?>("updateAll");

        return new UpdateFieldParameters(fieldIndex, updateAll);
    }

    /// <summary>
    ///     Parameters for the update field operation.
    /// </summary>
    /// <param name="FieldIndex">The index of the field to update, or null to update all.</param>
    /// <param name="UpdateAll">Whether to update all fields.</param>
    private sealed record UpdateFieldParameters(int? FieldIndex, bool? UpdateAll);
}
