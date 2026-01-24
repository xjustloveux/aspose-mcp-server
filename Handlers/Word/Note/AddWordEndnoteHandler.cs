using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for adding endnotes to Word documents.
/// </summary>
public class AddWordEndnoteHandler : AddWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "add_endnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Endnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Endnote";
}
