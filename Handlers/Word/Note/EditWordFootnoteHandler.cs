using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for editing footnotes in Word documents.
/// </summary>
public class EditWordFootnoteHandler : EditWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "edit_footnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Footnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Footnote";
}
