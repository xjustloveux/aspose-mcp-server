using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for editing endnotes in Word documents.
/// </summary>
public class EditWordEndnoteHandler : EditWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "edit_endnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Endnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Endnote";
}
