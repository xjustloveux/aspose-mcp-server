using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     What one refresh of a document's fields did.
/// </summary>
/// <param name="Updated">Fields that were resolved.</param>
/// <param name="Refused">Fields left alone because their type reaches outside the document.</param>
/// <param name="Locked">Fields left alone because the document locks them.</param>
public readonly record struct FieldUpdateTally(int Updated, int Refused, int Locked);

/// <summary>
///     Decides which Word fields this server is willing to create or resolve.
///     <para>
///         A Word field is a small program: updating an <c>INCLUDETEXT</c> field reads whatever
///         path it names, and <c>INCLUDEPICTURE</c> reads whatever URL it names. Measured against
///         the pinned Aspose.Words, an <c>INCLUDETEXT</c> field resolves on
///         <see cref="Field.Update()" /> and on <see cref="Document.UpdateFields" /> — so a field
///         code supplied by a caller is a file read or an outbound request, with the path sitting
///         somewhere the allowlist never inspects.
///     </para>
///     <para>
///         The header and footer writer was restricted first, but the general <c>word_field</c>
///         entry points inserted an arbitrary code and updated it immediately, which was simply a
///         second door to the same room (R3-S01). This type is the single place that decides, so a
///         new entry point cannot quietly reopen it.
///     </para>
/// </summary>
public static class WordFieldPolicy
{
    /// <summary>
    ///     Fields this server will create and resolve: everything here is computed from the
    ///     document or its metadata and reaches neither the filesystem nor the network.
    /// </summary>
    private static readonly HashSet<FieldType> Allowed =
    [
        FieldType.FieldPage,
        FieldType.FieldNumPages,
        FieldType.FieldNumChars,
        FieldType.FieldNumWords,
        FieldType.FieldDate,
        FieldType.FieldTime,
        FieldType.FieldCreateDate,
        FieldType.FieldSaveDate,
        FieldType.FieldPrintDate,
        FieldType.FieldEditTime,
        FieldType.FieldRevisionNum,
        FieldType.FieldFileName,
        FieldType.FieldFileSize,
        FieldType.FieldAuthor,
        FieldType.FieldLastSavedBy,
        FieldType.FieldTitle,
        FieldType.FieldSubject,
        FieldType.FieldKeyword,
        FieldType.FieldComments,
        FieldType.FieldTemplate,
        FieldType.FieldSectionPages,
        FieldType.FieldSection,
        FieldType.FieldTOC,
        FieldType.FieldTOCEntry,
        FieldType.FieldIndex,
        FieldType.FieldIndexEntry,
        FieldType.FieldRef,
        FieldType.FieldPageRef,
        FieldType.FieldSequence,
        FieldType.FieldStyleRef,
        FieldType.FieldListNum,
        FieldType.FieldNoteRef,
        FieldType.FieldFormTextInput,
        FieldType.FieldFormCheckBox,
        FieldType.FieldFormDropDown,
        FieldType.FieldMergeField,
        FieldType.FieldHyperlink
    ];

    /// <summary>
    ///     The names accepted where a caller supplies a field type as text, derived from
    ///     <see cref="Allowed" /> so the two can never disagree.
    /// </summary>
    private static readonly HashSet<string> AllowedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "PAGE", "NUMPAGES", "NUMCHARS", "NUMWORDS", "DATE", "TIME", "CREATEDATE", "SAVEDATE",
        "PRINTDATE", "EDITTIME", "REVNUM", "FILENAME", "FILESIZE", "AUTHOR", "LASTSAVEDBY",
        "TITLE", "SUBJECT", "KEYWORDS", "COMMENTS", "TEMPLATE", "SECTIONPAGES", "SECTION",
        "TOC", "TC", "INDEX", "XE", "REF", "PAGEREF", "SEQ", "STYLEREF", "LISTNUM", "SEQCAPTION",
        "NOTEREF", "FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN", "MERGEFIELD", "HYPERLINK"
    };

    /// <summary>Field names this server refuses outright, named so the error can be specific.</summary>
    private static readonly HashSet<string> KnownDangerousNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "INCLUDETEXT", "INCLUDEPICTURE", "INCLUDE", "DDE", "DDEAUTO", "LINK", "IMPORT",
        "AUTOTEXT", "AUTOTEXTLIST", "DATABASE", "PRINT", "MACROBUTTON", "GOTOBUTTON"
    };

    /// <summary>Whether an existing field may be resolved.</summary>
    /// <param name="type">The field's type.</param>
    /// <returns><c>true</c> when the field computes its result without external input.</returns>
    public static bool IsAllowed(FieldType type)
    {
        return Allowed.Contains(type);
    }

    /// <summary>
    ///     Rejects a caller-supplied field type before anything is inserted.
    /// </summary>
    /// <param name="fieldType">Field type name as the caller wrote it, e.g. <c>DATE</c>.</param>
    /// <exception cref="ArgumentException">Thrown when the field is not one this server creates.</exception>
    public static void EnsureAllowedName(string? fieldType)
    {
        var name = (fieldType ?? string.Empty).Trim();
        if (AllowedNames.Contains(name)) return;

        var reason = KnownDangerousNames.Contains(name)
            ? $"Field '{name}' resolves content from a file, another document or an external "
              + "program, so it is not available through this server."
            : $"Unsupported field type '{name}'.";

        throw new ArgumentException(
            reason + " Supported: " + string.Join(", ", AllowedNames.Order(StringComparer.Ordinal)),
            nameof(fieldType));
    }

    /// <summary>
    ///     Rejects a caller-supplied field <em>code</em> before anything is written.
    ///     <para>
    ///         <see cref="EnsureAllowedName" /> guards the add path, where the caller names a field
    ///         type. The edit path takes a whole field code instead and went straight to the
    ///         document: the existing runs were removed and the caller's text written in their
    ///         place, so <c>INCLUDETEXT "C:\secrets.txt"</c> became a live field that a later
    ///         update — including this server's own document-wide refresh — would resolve. The type
    ///         that was checked afterwards was the <em>old</em> field's, which the edit had already
    ///         replaced (R8-W01).
    ///     </para>
    ///     <para>
    ///         Both the leading command and any command written inside a nested <c>{ }</c> are
    ///         checked, so a permitted outer command cannot carry a refused one.
    ///     </para>
    /// </summary>
    /// <param name="fieldCode">The field code as the caller wrote it, e.g. <c>PAGE</c>.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the code names a field this server does not create.
    /// </exception>
    public static void EnsureAllowedFieldCode(string? fieldCode)
    {
        var code = (fieldCode ?? string.Empty).Trim();
        if (code.Length == 0) return;

        foreach (var command in CommandsIn(code))
            EnsureAllowedName(command);
    }

    /// <summary>
    ///     Every field command a field code names: the one it starts with, and one for each nested
    ///     field marker inside it.
    /// </summary>
    /// <param name="code">The trimmed field code.</param>
    /// <returns>The command names, in the order they appear.</returns>
    private static IEnumerable<string> CommandsIn(string code)
    {
        var outer = code.StartsWith('{') ? code[1..] : code;
        var leading = LeadingCommand(outer);
        if (leading.Length > 0) yield return leading;

        for (var i = code.IndexOf('{'); i >= 0; i = code.IndexOf('{', i + 1))
        {
            if (i == 0 && !ReferenceEquals(outer, code)) continue;

            var nested = LeadingCommand(code[(i + 1)..]);
            if (nested.Length > 0) yield return nested;
        }
    }

    /// <summary>The first whitespace-delimited token of a field code.</summary>
    /// <param name="text">Text beginning at a field code.</param>
    /// <returns>The command name, or an empty string when there is none.</returns>
    private static string LeadingCommand(string text)
    {
        var trimmed = text.TrimStart();
        var end = trimmed.IndexOfAny([' ', '\t', '\r', '\n', '}', '\\']);
        return (end < 0 ? trimmed : trimmed[..end]).Trim();
    }

    /// <summary>
    ///     Refuses any field the policy does not allow, at the moment the library is about to
    ///     resolve it.
    ///     <para>
    ///         Checking a field's own type before calling <see cref="Field.Update()" /> only
    ///         answers for the field named. A field's code can contain another field, and updating
    ///         the outer one updates what is inside it: measured against the pinned Aspose.Words,
    ///         a <c>REF</c>, <c>HYPERLINK</c> or <c>TOC</c> field carrying a nested
    ///         <c>INCLUDETEXT</c> read the file that field named, while the caller was told the
    ///         dangerous field had been skipped (R4-S01).
    ///     </para>
    ///     <para>
    ///         The library consults this for every field it resolves, nested ones included, so it
    ///         is the only place the answer has to be given.
    ///     </para>
    /// </summary>
    /// <summary>
    ///     The one refusal message, so a refusal raised before the update reads exactly like one
    ///     raised by the callback during it.
    /// </summary>
    /// <param name="type">The field type being refused.</param>
    /// <returns>The message.</returns>
    private static string DisallowedFieldMessage(FieldType type)
    {
        return $"A {type} field would be resolved as part of this update. That field reads "
               + "content from a file, another document or an external program, so it is not "
               + "available through this server. Remove it, or update a field that does not "
               + "contain it.";
    }

    /// <summary>
    ///     Installs the refusal callback for the duration of an update, and puts back whatever was
    ///     there before.
    /// </summary>
    /// <param name="document">The document about to have fields updated.</param>
    /// <returns>A scope that restores the document's previous callback when disposed.</returns>
    public static IDisposable GuardFieldUpdates(Document document)
    {
        return new UpdateGuard(document);
    }

    /// <summary>
    ///     Updates one field, refusing it and anything nested inside it that the policy denies.
    /// </summary>
    /// <param name="field">The field to update.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the field, or a field within it, is one this server does not resolve.
    /// </exception>
    public static void UpdateField(Field field)
    {
        if (!IsAllowed(field.Type))
            throw new ArgumentException(
                $"Field type {field.Type} resolves content from a file, another document or an "
                + "external program, so it is not available through this server.");

        var document = (Document)field.Start.Document;

        // The same read-only pass update-all does, for the one field being updated. Checking only
        // this field's own type left a nested refusal to be raised from inside Aspose's traversal,
        // after the outer field's own nested fields had already been resolved (R7-W01).
        RefuseNestedDisallowedFields(document, [field]);

        using var guard = GuardFieldUpdates(document);
        field.Update();
    }

    /// <summary>
    ///     Updates every field the policy allows, and leaves the rest untouched.
    ///     <para>
    ///         This replaces <see cref="Document.UpdateFields" />, which resolves everything in the
    ///         document including fields that arrived in the caller's own input file. Updating one
    ///         by one keeps the intended refresh while never resolving a field that reads a path.
    ///     </para>
    ///     <para>
    ///         The per-field type check is kept because it decides what to skip quietly; the
    ///         callback is what makes the decision hold for fields nested inside an allowed one
    ///         (R4-S01).
    ///     </para>
    /// </summary>
    /// <param name="document">The document whose fields should be refreshed.</param>
    /// <returns>
    ///     How many fields were updated, how many were left because the policy refuses their type,
    ///     and how many were left because the document locks them. Each field falls in exactly one
    ///     of the three, so the caller can report them without arithmetic that can go negative
    ///     (R4-C01).
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when an allowed field contains a field this server does not resolve.
    /// </exception>
    public static FieldUpdateTally UpdateAllowedFields(Document document)
    {
        // Read the whole document before changing any of it. The callback below is the guard that
        // cannot be fooled, but it fires part-way through Aspose's own traversal: a document whose
        // first field was allowed and changed, and whose second contained a refused field, threw
        // with the first field's new result already in place — and a session keeps that document
        // (R5-T01). Anything that would be refused is found first, so the refusal costs nothing.
        RefuseNestedDisallowedFields(document, null);

        using var guard = GuardFieldUpdates(document);

        var updated = 0;
        var refused = 0;
        var locked = 0;

        foreach (var field in document.Range.Fields)
        {
            // Ordered so each field is counted once: a locked field the policy also refuses is
            // reported as refused, because that is the reason that would stand even if it were
            // unlocked.
            if (!IsAllowed(field.Type))
            {
                refused++;
                continue;
            }

            if (field.IsLocked)
            {
                locked++;
                continue;
            }

            field.Update();
            updated++;
        }

        return new FieldUpdateTally(updated, refused, locked);
    }

    /// <summary>
    ///     Refuses before any update when a field that is about to be updated contains one this
    ///     server would not resolve.
    ///     <para>
    ///         Every entry point that updates a field goes through this first — the single-field
    ///         path, update-all, and the table-of-contents handler — because a refusal raised from
    ///         inside Aspose's traversal leaves whatever it had already updated in place (R7-W01).
    ///     </para>
    /// </summary>
    /// <param name="document">The document about to be updated.</param>
    /// <param name="fieldsToUpdate">
    ///     The fields the caller intends to update, or <c>null</c> for every field this server
    ///     would update in the whole document.
    /// </param>
    /// <param name="lockStateAfterwards">
    ///     Whether a field is locked once the caller's request has been applied. Omitted, each
    ///     field's current lock state is used — correct for a caller that does not change it, and
    ///     wrong for one that unlocks a field and updates it in the same request (R9-W01).
    /// </param>
    /// <exception cref="ArgumentException">
    ///     Thrown when an update would resolve a disallowed field. The document is untouched.
    /// </exception>
    internal static void RefuseNestedDisallowedFields(Document document,
        IReadOnlyCollection<Field>? fieldsToUpdate,
        Func<Field, bool>? lockStateAfterwards = null)
    {
        var fields = document.Range.Fields.ToList();
        var disallowed = fields.Where(field => !IsAllowed(field.Type)).ToList();
        if (disallowed.Count == 0) return;

        // A field this server will not update resolves nothing, and a locked one is skipped, so
        // neither can carry a refusal into the document.
        //
        // "Locked" means locked once the caller's request has been applied, which is not always
        // the state the field is in right now: a request that unlocks a field and updates it in
        // one go made the field look skippable at the moment it was examined and updatable a
        // moment later, and the nested field it resolved was never refused (R9-W01). Callers that
        // change lock state say so; the rest read it as it stands.
        var locked = lockStateAfterwards ?? (field => field.IsLocked);
        var updating = (fieldsToUpdate ?? fields)
            .Where(field => IsAllowed(field.Type) && !locked(field))
            .ToList();
        if (updating.Count == 0) return;

        // By document position, not by sibling walk: a field whose markers straddle a paragraph
        // break has them in different parents, and the walk simply never reached the end (R7-W01).
        var order = FieldBoundaryHelper.DocumentOrder(document);

        // One sweep, not every pair. The nested loop was O(updating × disallowed) with both counts
        // chosen by the document, which let a document keep the server busy before any field
        // was touched (R21-RES03). A field with a missing marker has no interval and cannot
        // contain or be contained, which is what the pairwise check also concluded.
        var outers = Intervals(updating, order);
        var inners = Intervals(disallowed, order);

        var contained = FieldBoundaryHelper.FirstContained(outers.Spans, inners.Spans);
        if (contained >= 0)
            throw new ArgumentException(DisallowedFieldMessage(inners.Fields[contained].Type));
    }

    /// <summary>The document-order span of every field that has both markers.</summary>
    /// <param name="fields">The fields.</param>
    /// <param name="order">Document order.</param>
    /// <returns>The spans and, index-aligned, the fields they belong to.</returns>
    private static (List<(int Start, int End)> Spans, List<Field> Fields) Intervals(
        IEnumerable<Field> fields, Dictionary<Node, int> order)
    {
        var spans = new List<(int Start, int End)>();
        var owners = new List<Field>();

        foreach (var field in fields)
        {
            if (field.Start == null || field.End == null) continue;
            if (!order.TryGetValue(field.Start, out var start)) continue;
            if (!order.TryGetValue(field.End, out var end)) continue;

            spans.Add((start, end));
            owners.Add(field);
        }

        return (spans, owners);
    }

    private sealed class RefuseDisallowedFields : IFieldUpdatingCallback
    {
        /// <summary>Refuses the field if the policy does not allow it.</summary>
        /// <param name="field">The field the library is about to resolve.</param>
        /// <exception cref="ArgumentException">Thrown when the field is not one this server resolves.</exception>
        public void FieldUpdating(Field field)
        {
            if (IsAllowed(field.Type)) return;

            throw new ArgumentException(DisallowedFieldMessage(field.Type));
        }

        /// <inheritdoc />
        public void FieldUpdated(Field field)
        {
        }
    }

    /// <summary>
    ///     Restores a document's previous field-updating callback when the update finishes.
    /// </summary>
    private sealed class UpdateGuard : IDisposable
    {
        private readonly Document _document;
        private readonly IFieldUpdatingCallback? _previous;

        /// <summary>Installs the refusal callback.</summary>
        /// <param name="document">The document to guard.</param>
        public UpdateGuard(Document document)
        {
            _document = document;
            _previous = document.FieldOptions.FieldUpdatingCallback;
            document.FieldOptions.FieldUpdatingCallback = new RefuseDisallowedFields();
        }

        /// <inheritdoc />
        public void Dispose()
        {
            _document.FieldOptions.FieldUpdatingCallback = _previous;
        }
    }
}
