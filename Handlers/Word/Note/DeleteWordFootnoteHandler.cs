using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for deleting footnotes from Word documents.
/// </summary>
public class DeleteWordFootnoteHandler : DeleteWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "delete_footnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Footnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Footnote";
}
