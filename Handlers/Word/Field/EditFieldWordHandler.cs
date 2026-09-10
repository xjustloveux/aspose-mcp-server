using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for editing fields in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits a field's code, lock state, or triggers an update.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldIndex
    ///     Optional: fieldCode, lockField, unlockField, updateField
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditFieldParameters(parameters);

        var document = context.Document;
        var fields = document.Range.Fields.ToList();

        if (p.FieldIndex < 0 || p.FieldIndex >= fields.Count)
            throw new ArgumentException(
                $"Field index {p.FieldIndex} is out of range (document has {fields.Count} fields)");

        var field = fields[p.FieldIndex];

        // Deny-by-default before a single node is touched. The type checked further down is the
        // one the field had *before* the edit, so a refused command written here would already be
        // in the document by the time anything looked at it (R8-W01).
        WordFieldPolicy.EnsureAllowedFieldCode(p.FieldCode);

        // And the document-wide refusal too, before the code and the lock state are changed: the
        // update below can refuse because of a field somewhere else entirely, which used to leave
        // this field already edited (R8-W02).
        // Judged on the state this request leaves behind, not the state it finds: a request that
        // unlocks the field and updates it in one go used to slip past entirely (R9-W01).
        if (p.UpdateFieldAfter)
            RefuseIfTheRequestWouldResolveADisallowedField(document, field, p.FieldIndex, p);

        var oldFieldCode = field.GetFieldCode();
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.FieldCode))
            UpdateFieldCode(document, field, p.FieldCode, oldFieldCode, changes);

        ApplyLockState(field, p.LockField, p.UnlockField, changes);

        if (p.UpdateFieldAfter)
        {
            // Both the edited field and the document-wide refresh go through the policy: the
            // document may already contain fields that arrived in the caller's own input. The
            // preflight above judged the same thing on the state this request leaves behind, so a
            // refusal here would be one it already made — this is the belt to that preflight's
            // braces, and it must never be the first place a request is stopped (R9-W01).
            if (!WordFieldPolicy.IsAllowed(field.Type))
                throw new ArgumentException(
                    $"Field #{p.FieldIndex} is a {field.Type} field, which this server does not "
                    + "resolve because it can read a file or fetch a URL.");

            WordFieldPolicy.UpdateField(field);
            WordFieldPolicy.UpdateAllowedFields(document);
        }

        MarkModified(context);

        var message = $"Field #{p.FieldIndex} edited successfully\n";
        message += $"Original field code: {oldFieldCode}\n";
        if (changes.Count > 0)
            message += $"Changes: {string.Join(", ", changes)}";
        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Refuses, before anything is written, if the request as a whole would resolve a field
    ///     this server does not resolve.
    ///     <para>
    ///         The lock state the check uses is the one the request <em>leaves behind</em>, not
    ///         the one it finds. A locked field is skipped by the nested-field check because a
    ///         locked field resolves nothing — which stops being true the moment the same request
    ///         unlocks it. Measured, <c>unlockField</c> with <c>updateField</c> resolved a nested
    ///         INCLUDETEXT with no refusal at all, and the same request carrying a field code
    ///         refused only after the code and the lock had been written (R9-W01).
    ///     </para>
    ///     <para>
    ///         The field's own eligibility is judged the same way. A caller-supplied code has
    ///         already been through <see cref="WordFieldPolicy.EnsureAllowedFieldCode" />, so a
    ///         request that carries one leaves an allowed field behind whatever the field was
    ///         before; a request that carries none leaves the field's current type.
    ///     </para>
    /// </summary>
    /// <param name="document">The session's document.</param>
    /// <param name="field">The field the request edits.</param>
    /// <param name="index">That field's index, for the message.</param>
    /// <param name="p">The request.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the request would update a disallowed field, or would resolve one nested
    ///     inside a field it updates.
    /// </exception>
    private static void RefuseIfTheRequestWouldResolveADisallowedField(Document document,
        Aspose.Words.Fields.Field field, int index, EditFieldParameters p)
    {
        var replacingTheCode = !string.IsNullOrEmpty(p.FieldCode);
        if (!replacingTheCode && !WordFieldPolicy.IsAllowed(field.Type))
            throw new ArgumentException(
                $"Field #{index} is a {field.Type} field, which this server does not "
                + "resolve because it can read a file or fetch a URL.");

        // Matched on the start marker, not on the wrapper: the policy enumerates the document's
        // fields itself, so the Field object it hands back is a different instance describing the
        // same field. A reference comparison here silently matched nothing.
        WordFieldPolicy.RefuseNestedDisallowedFields(document, null,
            candidate => ReferenceEquals(candidate.Start, field.Start)
                ? WillBeLocked(field, p)
                : candidate.IsLocked);
    }

    /// <summary>
    ///     Whether the edited field is locked once this request has been applied.
    /// </summary>
    /// <param name="field">The field being edited.</param>
    /// <param name="p">The request.</param>
    /// <returns>The lock state the request leaves behind.</returns>
    private static bool WillBeLocked(Aspose.Words.Fields.Field field, EditFieldParameters p)
    {
        if (p.LockField == true) return true;
        if (p.UnlockField == true) return false;
        return field.IsLocked;
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

        // Only the code region is replaced: the runs between the start marker and the separator.
        // The result region, after the separator, is what the field currently displays and is left
        // for the update to refresh.
        var currentNode = fieldStart.NextSibling;
        var stopAt = (Node?)field.Separator ?? fieldEnd;
        while (currentNode != null && currentNode != stopAt)
        {
            var nextNode = currentNode.NextSibling;
            currentNode.Remove();
            currentNode = nextNode;
        }

        // Moving to the start marker inserts *before* it, which put the new code in the body as
        // ordinary text and left the field itself empty: the operation reported "Field code
        // updated" while the document showed the code and the field had become FieldNone
        // (measured on the pinned Aspose.Words). Moving to the node the code region ends at puts
        // the text inside that region, which is where a field code lives.
        builder.MoveTo(stopAt);
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
            throw new ArgumentException("fieldIndex is required for edit operation");

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
