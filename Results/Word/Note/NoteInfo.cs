using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Note;

/// <summary>
///     Information about a single footnote or endnote.
/// </summary>
public sealed record NoteInfo
{
    /// <summary>
    ///     Zero-based index of the note.
    /// </summary>
    [JsonPropertyName("noteIndex")]
    public required int NoteIndex { get; init; }

    /// <summary>
    ///     The reference mark used for the note.
    /// </summary>
    [JsonPropertyName("referenceMark")]
    public required string ReferenceMark { get; init; }

    /// <summary>
    ///     The text content of the note.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }
}
