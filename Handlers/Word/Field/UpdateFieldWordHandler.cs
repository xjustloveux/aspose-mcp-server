using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for updating fields in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class UpdateFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "update";

    /// <summary>
    ///     Updates one or all fields in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: fieldIndex (int) — update only the field at this index (0-based).
    ///     Optional: updateAll (bool) — explicitly request updating all fields.
    ///     If neither parameter is provided, all fields are updated by default.
    ///     If both are provided, updateAll wins: asking for all of them is the broader request,
    ///     and the implementation has always taken it that way. The documentation used to claim
    ///     the opposite (R4-DOC03).
    /// </param>
    /// <returns>
    ///     Success message naming how many fields were updated, and how many were left because
    ///     the document locks them or because their type reaches outside the document.
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the named field is one this server does not resolve, or when an allowed
    ///     field contains one (R4-S01).
    /// </exception>
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

            if (!WordFieldPolicy.IsAllowed(field.Type))
                throw new ArgumentException(
                    $"Field #{p.FieldIndex.Value} is a {field.Type} field, which this server does "
                    + "not resolve because it can read a file or fetch a URL.");

            var oldResult = field.Result ?? "";
            WordFieldPolicy.UpdateField(field);
            var newResult = field.Result ?? "";

            // Reaching this line is itself the result: a locked field returned above with a warning
            // and a refused one threw, so the only way here is a field this call actually updated.
            // Comparing the old and new result text instead would call a real update a no-op —
            // measured: DocumentBuilder.InsertField evaluates the field as it inserts it, so
            // updating it again rewrites the same text (R5-C02).
            MarkModified(context);

            return new SuccessResult
                { Message = $"Field #{p.FieldIndex.Value} updated\nOld result: {oldResult}\nNew result: {newResult}" };
        }

        // Counted by the update itself rather than subtracted here: a locked field this server
        // refuses anyway was counted twice, so a document holding one reported "Updated -1"
        // (R4-C01).
        var tally = WordFieldPolicy.UpdateAllowedFields(document);

        // Marked unconditionally before: a document with no fields, or one whose every field was
        // locked or refused, reported "Updated 0" and still left the session dirty (R5-C02).
        if (tally.Updated > 0) MarkModified(context);

        var message = $"Updated {tally.Updated} field(s)";
        var lockedCount = tally.Locked;
        var refused = tally.Refused;
        if (lockedCount > 0)
            message += $"\nSkipped {lockedCount} locked field(s)";
        if (refused > 0)
            message += $"\nSkipped {refused} field(s) that resolve external content";
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
