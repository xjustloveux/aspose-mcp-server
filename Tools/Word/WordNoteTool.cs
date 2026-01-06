using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for footnote and endnote operations in Word documents
///     Merges: WordAddFootnoteTool, WordAddEndnoteTool, WordDeleteFootnoteTool, WordDeleteEndnoteTool,
///     WordEditFootnoteTool, WordEditEndnoteTool, WordGetFootnotesTool, WordGetEndnotesTool
/// </summary>
[McpServerToolType]
public class WordNoteTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordNoteTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordNoteTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word note operation (add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote,
    ///     edit_endnote, get_footnotes, get_endnotes).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_footnote, add_endnote, delete_footnote, delete_endnote,
    ///     edit_footnote, edit_endnote, get_footnotes, get_endnotes.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Note text content.</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for document end).</param>
    /// <param name="sectionIndex">Section index (0-based, default: 0).</param>
    /// <param name="referenceText">Reference text in document to insert note at.</param>
    /// <param name="customMark">Custom note mark.</param>
    /// <param name="referenceMark">Reference mark of note to delete/edit.</param>
    /// <param name="noteIndex">Note index (0-based).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_note")]
    [Description(
        @"Manage footnotes and endnotes in Word documents. Supports 8 operations: add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes.

Usage examples:
- Add footnote: word_note(operation='add_footnote', path='doc.docx', text='Footnote text', paragraphIndex=0)
- Add endnote: word_note(operation='add_endnote', path='doc.docx', text='Endnote text', paragraphIndex=0)
- Delete footnote: word_note(operation='delete_footnote', path='doc.docx', noteIndex=0)
- Edit footnote: word_note(operation='edit_footnote', path='doc.docx', noteIndex=0, text='Updated footnote')
- Get footnotes: word_note(operation='get_footnotes', path='doc.docx')")]
    public string Execute(
        [Description(
            "Operation: add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Note text content")] string? text = null,
        [Description("Paragraph index (0-based, -1 for document end)")]
        int? paragraphIndex = null,
        [Description("Section index (0-based, default: 0)")]
        int sectionIndex = 0,
        [Description("Reference text in document to insert note at")]
        string? referenceText = null,
        [Description("Custom note mark")] string? customMark = null,
        [Description("Reference mark of note to delete/edit")]
        string? referenceMark = null,
        [Description("Note index (0-based)")] int? noteIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add_footnote" => AddNote(ctx, outputPath, text, paragraphIndex, sectionIndex, referenceText, customMark,
                FootnoteType.Footnote),
            "add_endnote" => AddNote(ctx, outputPath, text, paragraphIndex, sectionIndex, referenceText, customMark,
                FootnoteType.Endnote),
            "delete_footnote" => DeleteNote(ctx, outputPath, referenceMark, noteIndex, FootnoteType.Footnote),
            "delete_endnote" => DeleteNote(ctx, outputPath, referenceMark, noteIndex, FootnoteType.Endnote),
            "edit_footnote" => EditNote(ctx, outputPath, referenceMark, noteIndex, text, FootnoteType.Footnote),
            "edit_endnote" => EditNote(ctx, outputPath, referenceMark, noteIndex, text, FootnoteType.Endnote),
            "get_footnotes" => GetNotes(ctx, FootnoteType.Footnote),
            "get_endnotes" => GetNotes(ctx, FootnoteType.Endnote),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets the display name for the note type.
    /// </summary>
    /// <param name="type">The footnote type.</param>
    /// <returns>The display name ("footnote" or "endnote").</returns>
    private static string GetNoteTypeName(FootnoteType type)
    {
        return type == FootnoteType.Footnote ? "footnote" : "endnote";
    }

    /// <summary>
    ///     Gets all notes of the specified type from the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="type">The footnote type to retrieve.</param>
    /// <returns>A list of footnotes of the specified type.</returns>
    private static List<Footnote> GetNotesFromDoc(Document doc, FootnoteType type)
    {
        return doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == type)
            .ToList();
    }

    /// <summary>
    ///     Adds a footnote or endnote to the document at the specified location.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The note text content.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index, or -1 for document end.</param>
    /// <param name="sectionIndex">The zero-based section index.</param>
    /// <param name="referenceText">The reference text in document to insert note at.</param>
    /// <param name="customMark">The custom note mark.</param>
    /// <param name="noteType">The type of note (footnote or endnote).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when text is empty, section/paragraph index is out of range, or reference
    ///     text is not found.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to add an endnote in a header or footer.</exception>
    private static string AddNote(DocumentContext<Document> ctx, string? outputPath, string? text, int? paragraphIndex,
        int sectionIndex, string? referenceText, string? customMark, FootnoteType noteType)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for add operation");

        var typeName = GetNoteTypeName(noteType);
        var doc = ctx.Document;

        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var builder = new DocumentBuilder(doc);
        Footnote? insertedNote;

        if (!string.IsNullOrEmpty(referenceText))
        {
            var callback = new NoteInsertingCallback(builder, noteType, text, customMark);
            var options = new FindReplaceOptions { ReplacingCallback = callback };
            doc.Range.Replace(referenceText, referenceText, options);

            insertedNote = callback.InsertedNote ??
                           throw new ArgumentException($"Reference text '{referenceText}' not found");
        }
        else if (paragraphIndex.HasValue)
        {
            var section = doc.Sections[sectionIndex];

            if (paragraphIndex.Value == -1)
            {
                var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();
                if (bodyParagraphs.Count > 0)
                    builder.MoveTo(bodyParagraphs[^1]);
                else
                    builder.MoveToDocumentEnd();
            }
            else
            {
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                    throw new ArgumentException(
                        $"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

                var para = paragraphs[paragraphIndex.Value];

                if (noteType == FootnoteType.Endnote)
                {
                    var parentNode = para.ParentNode;
                    while (parentNode != null)
                    {
                        if (parentNode is HeaderFooter)
                            throw new InvalidOperationException(
                                $"Endnotes are only allowed inside the main document body. The paragraph at index {paragraphIndex.Value} is located in a header or footer.");
                        if (parentNode is Section || parentNode is Body) break;
                        parentNode = parentNode.ParentNode;
                    }
                }

                builder.MoveTo(para);
            }

            insertedNote = builder.InsertFootnote(noteType, text);
            if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;
        }
        else
        {
            builder.MoveToDocumentEnd();
            insertedNote = builder.InsertFootnote(noteType, text);
            if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;
        }

        ctx.Save(outputPath);

        var result = new StringBuilder();
        result.AppendLine($"{char.ToUpper(typeName[0]) + typeName.Substring(1)} added successfully");
        result.AppendLine($"Text: {text}");
        if (insertedNote != null && !string.IsNullOrEmpty(insertedNote.ReferenceMark))
            result.AppendLine($"Reference mark: {insertedNote.ReferenceMark}");
        result.AppendLine(ctx.GetOutputMessage(outputPath));
        return result.ToString();
    }

    /// <summary>
    ///     Deletes a footnote or endnote from the document by reference mark or index.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="referenceMark">The reference mark of the note to delete.</param>
    /// <param name="noteIndex">The zero-based note index.</param>
    /// <param name="noteType">The type of note (footnote or endnote).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when note index is out of range.</exception>
    private static string DeleteNote(DocumentContext<Document> ctx, string? outputPath, string? referenceMark,
        int? noteIndex, FootnoteType noteType)
    {
        var typeName = GetNoteTypeName(noteType);
        var doc = ctx.Document;
        var notes = GetNotesFromDoc(doc, noteType);

        var deletedCount = 0;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            var note = notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
            if (note != null)
            {
                note.Remove();
                deletedCount = 1;
            }
        }
        else if (noteIndex.HasValue)
        {
            if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
            {
                notes[noteIndex.Value].Remove();
                deletedCount = 1;
            }
            else
            {
                throw new ArgumentException(
                    $"Note index {noteIndex.Value} is out of range (document has {notes.Count} {typeName}s, valid index: 0-{notes.Count - 1})");
            }
        }
        else
        {
            foreach (var note in notes)
            {
                note.Remove();
                deletedCount++;
            }
        }

        ctx.Save(outputPath);
        return $"Deleted {deletedCount} {typeName}(s)\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits the text content of a footnote or endnote.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="referenceMark">The reference mark of the note to edit.</param>
    /// <param name="noteIndex">The zero-based note index.</param>
    /// <param name="newText">The new text content for the note.</param>
    /// <param name="noteType">The type of note (footnote or endnote).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is empty or the specified note is not found.</exception>
    private static string EditNote(DocumentContext<Document> ctx, string? outputPath, string? referenceMark,
        int? noteIndex, string? newText, FootnoteType noteType)
    {
        if (string.IsNullOrEmpty(newText))
            throw new ArgumentException("text parameter is required for edit operation");

        var typeName = GetNoteTypeName(noteType);
        var doc = ctx.Document;
        var notes = GetNotesFromDoc(doc, noteType);

        Footnote? note = null;

        if (!string.IsNullOrEmpty(referenceMark))
        {
            note = notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);
        }
        else if (noteIndex.HasValue)
        {
            if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
                note = notes[noteIndex.Value];
        }
        else if (notes.Count > 0)
        {
            note = notes[0];
        }

        if (note == null)
        {
            var availableInfo = notes.Count > 0
                ? $" (document has {notes.Count} {typeName}s, valid index: 0-{notes.Count - 1})"
                : $" (document has no {typeName}s)";
            throw new ArgumentException(
                $"Specified {typeName} not found{availableInfo}. Use get_{typeName}s operation to view available {typeName}s");
        }

        var oldText = note.ToString(SaveFormat.Text).Trim();
        note.RemoveAllChildren();
        var para = new Paragraph(doc);
        note.AppendChild(para);
        var run = new Run(doc, newText);
        para.AppendChild(run);

        ctx.Save(outputPath);

        var result = new StringBuilder();
        result.AppendLine($"{char.ToUpper(typeName[0]) + typeName.Substring(1)} edited successfully");
        result.AppendLine($"Old text: {oldText}");
        result.AppendLine($"New text: {newText}");
        result.AppendLine(ctx.GetOutputMessage(outputPath));
        return result.ToString();
    }

    /// <summary>
    ///     Gets all footnotes or endnotes from the document as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="noteType">The type of note (footnote or endnote).</param>
    /// <returns>A JSON string containing the list of notes.</returns>
    private static string GetNotes(DocumentContext<Document> ctx, FootnoteType noteType)
    {
        var doc = ctx.Document;
        var typeName = GetNoteTypeName(noteType);
        var notes = GetNotesFromDoc(doc, noteType);

        List<object> noteList = [];
        for (var i = 0; i < notes.Count; i++)
        {
            var note = notes[i];
            noteList.Add(new
            {
                noteIndex = i,
                referenceMark = note.ReferenceMark,
                text = note.ToString(SaveFormat.Text).Trim()
            });
        }

        var result = new
        {
            noteType = typeName,
            count = notes.Count,
            notes = noteList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Callback handler for inserting notes at matched text locations.
    /// </summary>
    /// <param name="builder">The document builder for inserting notes.</param>
    /// <param name="noteType">The type of note (footnote or endnote).</param>
    /// <param name="noteText">The text content of the note.</param>
    /// <param name="customMark">Optional custom reference mark for the note.</param>
    private class NoteInsertingCallback(
        DocumentBuilder builder,
        FootnoteType noteType,
        string noteText,
        string? customMark) : IReplacingCallback
    {
        public Footnote? InsertedNote { get; private set; }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            var matchNode = args.MatchNode;
            builder.MoveTo(matchNode);

            InsertedNote = builder.InsertFootnote(noteType, noteText);
            if (!string.IsNullOrEmpty(customMark))
                InsertedNote.ReferenceMark = customMark;

            return ReplaceAction.Skip;
        }
    }
}