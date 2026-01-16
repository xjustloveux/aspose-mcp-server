using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for editing fields in Word documents.
/// </summary>
public class EditFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_field";

    /// <summary>
    ///     Edits a field's code, lock state, or triggers an update.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldIndex
    ///     Optional: fieldCode, lockField, unlockField, updateField
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditFieldParameters(parameters);

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (p.FieldIndex < 0 || p.FieldIndex >= fields.Count)
            throw new ArgumentException(
                $"Field index {p.FieldIndex} is out of range (document has {fields.Count} fields)");

        var field = fields[p.FieldIndex];
        var oldFieldCode = field.GetFieldCode();
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.FieldCode))
        {
            var fieldStart = field.Start;
            var fieldEnd = field.End;

            if (fieldStart != null && fieldEnd != null)
            {
                var builder = new DocumentBuilder(document);
                builder.MoveTo(fieldStart);

                var currentNode = fieldStart.NextSibling;
                while (currentNode != null && currentNode != fieldEnd)
                {
                    var nextNode = currentNode.NextSibling;
                    if (currentNode.NodeType != NodeType.FieldSeparator && currentNode.NodeType != NodeType.FieldEnd)
                        currentNode.Remove();
                    currentNode = nextNode;
                }

                builder.MoveTo(fieldStart);
                builder.Write(p.FieldCode);
                changes.Add($"Field code updated: {oldFieldCode} -> {p.FieldCode}");
            }
        }

        if (p.LockField == true)
        {
            field.IsLocked = true;
            changes.Add("Field locked");
        }
        else if (p.UnlockField == true)
        {
            field.IsLocked = false;
            changes.Add("Field unlocked");
        }

        if (p.UpdateFieldAfter)
        {
            field.Update();
            document.UpdateFields();
        }

        MarkModified(context);

        var result = $"Field #{p.FieldIndex} edited successfully\n";
        result += $"Original field code: {oldFieldCode}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}";
        return result;
    }

    /// <summary>
    ///     Extracts and validates parameters for the edit field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldIndex is not provided.</exception>
    private static EditFieldParameters ExtractEditFieldParameters(OperationParameters parameters)
    {
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var fieldCode = parameters.GetOptional<string?>("fieldCode");
        var lockField = parameters.GetOptional<bool?>("lockField");
        var unlockField = parameters.GetOptional<bool?>("unlockField");
        var updateFieldAfter = parameters.GetOptional("updateField", true);

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for edit_field operation");

        return new EditFieldParameters(fieldIndex.Value, fieldCode, lockField, unlockField, updateFieldAfter);
    }

    /// <summary>
    ///     Parameters for the edit field operation.
    /// </summary>
    /// <param name="FieldIndex">The index of the field to edit.</param>
    /// <param name="FieldCode">The new field code.</param>
    /// <param name="LockField">Whether to lock the field.</param>
    /// <param name="UnlockField">Whether to unlock the field.</param>
    /// <param name="UpdateFieldAfter">Whether to update the field after editing.</param>
    private record EditFieldParameters(
        int FieldIndex,
        string? FieldCode,
        bool? LockField,
        bool? UnlockField,
        bool UpdateFieldAfter);
}
