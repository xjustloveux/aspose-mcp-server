using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for deleting endnotes from Word documents.
/// </summary>
public class DeleteWordEndnoteHandler : DeleteWordNoteHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "delete_endnote";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Endnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Endnote";
}
