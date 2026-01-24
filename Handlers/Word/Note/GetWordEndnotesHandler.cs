using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for getting endnotes from Word documents.
/// </summary>
public class GetWordEndnotesHandler : GetWordNotesHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "get_endnotes";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Endnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Endnote";
}
