using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Information about a single paragraph.
/// </summary>
public record ParagraphInfo
{
    /// <summary>
    ///     Zero-based index of the paragraph.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Location type (Body, Comment, TextBox, Shape).
    /// </summary>
    [JsonPropertyName("location")]
    public required string Location { get; init; }

    /// <summary>
    ///     Style name.
    /// </summary>
    [JsonPropertyName("style")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Style { get; init; }

    /// <summary>
    ///     Paragraph text (truncated if too long).
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Full text length.
    /// </summary>
    [JsonPropertyName("textLength")]
    public required int TextLength { get; init; }

    /// <summary>
    ///     Comment information if in a comment.
    /// </summary>
    [JsonPropertyName("commentInfo")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CommentInfo { get; init; }
}
