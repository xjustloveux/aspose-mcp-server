using Aspose.Words.Notes;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Handler for getting footnotes from Word documents.
/// </summary>
public class ListWordFootnotesHandler : GetWordNotesHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "list_footnotes";

    /// <inheritdoc />
    protected override FootnoteType NoteType => FootnoteType.Footnote;

    /// <inheritdoc />
    protected override string NoteTypeName => "Footnote";
}
