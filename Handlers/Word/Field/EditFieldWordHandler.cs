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
            UpdateFieldCode(document, field, p.FieldCode, oldFieldCode, changes);

        ApplyLockState(field, p.LockField, p.UnlockField, changes);

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
    ///     Updates the field code by removing existing content and inserting new code.
    /// </summary>
    /// <param name="document">The document.</param>
    /// <param name="field">The field to update.</param>
    /// <param name="newFieldCode">The new field code.</param>
    /// <param name="oldFieldCode">The old field code for change tracking.</param>
    /// <param name="changes">The list of changes to record.</param>
    private static void UpdateFieldCode(Document document, Aspose.Words.Fields.Field field, string newFieldCode,
        string oldFieldCode, List<string> changes)
    {
        var fieldStart = field.Start;
        var fieldEnd = field.End;

        if (fieldStart == null || fieldEnd == null) return;

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
        builder.Write(newFieldCode);
        changes.Add($"Field code updated: {oldFieldCode} -> {newFieldCode}");
    }

    /// <summary>
    ///     Applies lock or unlock state to the field.
    /// </summary>
    /// <param name="field">The field to modify.</param>
    /// <param name="lockField">Whether to lock the field.</param>
    /// <param name="unlockField">Whether to unlock the field.</param>
    /// <param name="changes">The list of changes to record.</param>
    private static void ApplyLockState(Aspose.Words.Fields.Field field, bool? lockField, bool? unlockField,
        List<string> changes)
    {
        if (lockField == true)
        {
            field.IsLocked = true;
            changes.Add("Field locked");
        }
        else if (unlockField == true)
        {
            field.IsLocked = false;
            changes.Add("Field unlocked");
        }
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
    private sealed record EditFieldParameters(
        int FieldIndex,
        string? FieldCode,
        bool? LockField,
        bool? UnlockField,
        bool UpdateFieldAfter);
}
