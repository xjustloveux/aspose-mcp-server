using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for adding footnotes to Word documents.
/// </summary>
public class AddWordFootnoteHandler : AddWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "add_footnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Footnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Footnote";
}
