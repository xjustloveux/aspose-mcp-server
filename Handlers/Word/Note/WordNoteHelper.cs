using Aspose.Words;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Helper class for Word footnote and endnote operations.
/// </summary>
public static class WordNoteHelper
{
    /// <summary>
    ///     Gets the display name for the note type.
    /// </summary>
    /// <param name="type">The footnote type.</param>
    /// <returns>The display name ("footnote" or "endnote").</returns>
    public static string GetNoteTypeName(FootnoteType type)
    {
        return type == FootnoteType.Footnote ? "footnote" : "endnote";
    }

    /// <summary>
    ///     Gets all notes of the specified type from the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="type">The footnote type to retrieve.</param>
    /// <returns>A list of footnotes of the specified type.</returns>
    public static List<Footnote> GetNotesFromDoc(Document doc, FootnoteType type)
    {
        return doc.GetChildNodes(NodeType.Footnote, true).Cast<Footnote>()
            .Where(f => f.FootnoteType == type)
            .ToList();
    }

    /// <summary>
    ///     Inserts a note at the specified reference text location.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="builder">The document builder.</param>
    /// <param name="referenceText">The reference text to search for.</param>
    /// <param name="noteType">The type of note to insert.</param>
    /// <param name="noteText">The text content of the note.</param>
    /// <param name="customMark">Optional custom reference mark.</param>
    /// <returns>The inserted footnote.</returns>
    /// <exception cref="ArgumentException">Thrown when reference text is not found.</exception>
    public static Footnote InsertNoteAtReferenceText(Document doc, DocumentBuilder builder, string referenceText,
        FootnoteType noteType, string noteText, string? customMark)
    {
        var callback = new NoteInsertingCallback(builder, noteType, noteText, customMark);
        var options = new FindReplaceOptions { ReplacingCallback = callback };
        doc.Range.Replace(referenceText, referenceText, options);

        return callback.InsertedNote ??
               throw new ArgumentException($"Reference text '{referenceText}' not found");
    }

    /// <summary>
    ///     Inserts a note at the specified paragraph.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="section">The section containing the paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="noteType">The type of note to insert.</param>
    /// <param name="noteText">The text content of the note.</param>
    /// <param name="customMark">Optional custom reference mark.</param>
    /// <returns>The inserted footnote.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to add an endnote in a header or footer.</exception>
    public static Footnote InsertNoteAtParagraph(DocumentBuilder builder, Section section,
        int paragraphIndex, FootnoteType noteType, string noteText, string? customMark)
    {
        if (paragraphIndex == -1)
        {
            var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<WordParagraph>().ToList();
            if (bodyParagraphs.Count > 0)
                builder.MoveTo(bodyParagraphs[^1]);
            else
                builder.MoveToDocumentEnd();
        }
        else
        {
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"paragraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");

            var para = paragraphs[paragraphIndex];

            if (noteType == FootnoteType.Endnote)
            {
                var parentNode = para.ParentNode;
                while (parentNode != null)
                {
                    if (parentNode is Aspose.Words.HeaderFooter)
                        throw new InvalidOperationException(
                            $"Endnotes are only allowed inside the main document body. The paragraph at index {paragraphIndex} is located in a header or footer.");
                    if (parentNode is Section || parentNode is Body) break;
                    parentNode = parentNode.ParentNode;
                }
            }

            builder.MoveTo(para);
        }

        var insertedNote = builder.InsertFootnote(noteType, noteText);
        if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;

        return insertedNote;
    }

    /// <summary>
    ///     Inserts a note at the document end.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="noteType">The type of note to insert.</param>
    /// <param name="noteText">The text content of the note.</param>
    /// <param name="customMark">Optional custom reference mark.</param>
    /// <returns>The inserted footnote.</returns>
    public static Footnote InsertNoteAtDocumentEnd(DocumentBuilder builder, FootnoteType noteType,
        string noteText, string? customMark)
    {
        builder.MoveToDocumentEnd();
        var insertedNote = builder.InsertFootnote(noteType, noteText);
        if (!string.IsNullOrEmpty(customMark)) insertedNote.ReferenceMark = customMark;

        return insertedNote;
    }

    /// <summary>
    ///     Finds a note by reference mark or index.
    /// </summary>
    /// <param name="notes">The list of notes to search.</param>
    /// <param name="referenceMark">Optional reference mark to search for.</param>
    /// <param name="noteIndex">Optional note index.</param>
    /// <returns>The found note, or null if not found.</returns>
    public static Footnote? FindNote(List<Footnote> notes, string? referenceMark, int? noteIndex)
    {
        if (!string.IsNullOrEmpty(referenceMark))
            return notes.FirstOrDefault(f => f.ReferenceMark == referenceMark);

        if (noteIndex.HasValue)
        {
            // If an explicit index is provided, only return note if index is valid
            if (noteIndex.Value >= 0 && noteIndex.Value < notes.Count)
                return notes[noteIndex.Value];
            return null; // Invalid index should not fallback to first note
        }

        // Only return first note when no index was explicitly provided
        return notes.Count > 0 ? notes[0] : null;
    }

    /// <summary>
    ///     Updates the text content of a note.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="note">The note to update.</param>
    /// <param name="newText">The new text content.</param>
    public static void UpdateNoteText(Document doc, Footnote note, string newText)
    {
        note.RemoveAllChildren();
        var para = new WordParagraph(doc);
        note.AppendChild(para);
        var run = new Run(doc, newText);
        para.AppendChild(run);
    }

    /// <summary>
    ///     Callback handler for inserting notes at matched text locations.
    /// </summary>
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
