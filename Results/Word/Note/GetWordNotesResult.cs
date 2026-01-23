using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Note;

/// <summary>
///     Result for getting footnotes or endnotes from Word documents.
/// </summary>
public sealed record GetWordNotesResult
{
    /// <summary>
    ///     Type of notes (footnote or endnote).
    /// </summary>
    [JsonPropertyName("noteType")]
    public required string NoteType { get; init; }

    /// <summary>
    ///     Total number of notes.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of notes.
    /// </summary>
    [JsonPropertyName("notes")]
    public required IReadOnlyList<NoteInfo> Notes { get; init; }
}
