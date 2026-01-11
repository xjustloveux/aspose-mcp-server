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
        var fieldIndex = parameters.GetOptional<int?>("fieldIndex");
        var fieldCode = parameters.GetOptional<string?>("fieldCode");
        var lockField = parameters.GetOptional<bool?>("lockField");
        var unlockField = parameters.GetOptional<bool?>("unlockField");
        var updateFieldAfter = parameters.GetOptional("updateField", true);

        if (!fieldIndex.HasValue)
            throw new ArgumentException("fieldIndex is required for edit_field operation");

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
            throw new ArgumentException(
                $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

        var field = fields[fieldIndex.Value];
        var oldFieldCode = field.GetFieldCode();
        List<string> changes = [];

        if (!string.IsNullOrEmpty(fieldCode))
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
                builder.Write(fieldCode);
                changes.Add($"Field code updated: {oldFieldCode} -> {fieldCode}");
            }
        }

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

        if (updateFieldAfter)
        {
            field.Update();
            document.UpdateFields();
        }

        MarkModified(context);

        var result = $"Field #{fieldIndex.Value} edited successfully\n";
        result += $"Original field code: {oldFieldCode}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}";
        return result;
    }
}
