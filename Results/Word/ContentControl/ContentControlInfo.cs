using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.ContentControl;

/// <summary>
///     Information about a single content control (structured document tag).
/// </summary>
public record ContentControlInfo
{
    /// <summary>
    ///     Zero-based index of the content control in the document.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     The tag identifier of the content control.
    /// </summary>
    [JsonPropertyName("tag")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Tag { get; init; }

    /// <summary>
    ///     The title (display name) of the content control.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     The type of the content control (e.g., PlainText, RichText, DropDownList, DatePicker, CheckBox).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     The current text value of the content control.
    /// </summary>
    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }

    /// <summary>
    ///     The placeholder text of the content control.
    /// </summary>
    [JsonPropertyName("placeholder")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Placeholder { get; init; }

    /// <summary>
    ///     Whether the content control contents are locked from editing.
    /// </summary>
    [JsonPropertyName("lockContents")]
    public bool LockContents { get; init; }

    /// <summary>
    ///     Whether the content control is locked from deletion.
    /// </summary>
    [JsonPropertyName("lockDeletion")]
    public bool LockDeletion { get; init; }
}
