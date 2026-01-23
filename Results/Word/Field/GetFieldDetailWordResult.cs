using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Result for getting detailed information about a specific field.
/// </summary>
public sealed record GetFieldDetailWordResult
{
    /// <summary>
    ///     Zero-based index of the field.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Field type name (Date, PageRef, Hyperlink, etc.).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Field type code (integer value of the enum).
    /// </summary>
    [JsonPropertyName("typeCode")]
    public required int TypeCode { get; init; }

    /// <summary>
    ///     Field code.
    /// </summary>
    [JsonPropertyName("code")]
    public required string Code { get; init; }

    /// <summary>
    ///     Field result/display value.
    /// </summary>
    [JsonPropertyName("result")]
    public string? Result { get; init; }

    /// <summary>
    ///     Indicates whether the field is locked.
    /// </summary>
    [JsonPropertyName("isLocked")]
    public required bool IsLocked { get; init; }

    /// <summary>
    ///     Indicates whether the field result is dirty (needs update).
    /// </summary>
    [JsonPropertyName("isDirty")]
    public required bool IsDirty { get; init; }

    /// <summary>
    ///     Hyperlink address (for hyperlink fields).
    /// </summary>
    [JsonPropertyName("hyperlinkAddress")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HyperlinkAddress { get; init; }

    /// <summary>
    ///     Hyperlink screen tip (for hyperlink fields).
    /// </summary>
    [JsonPropertyName("hyperlinkScreenTip")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HyperlinkScreenTip { get; init; }

    /// <summary>
    ///     Bookmark name (for reference fields).
    /// </summary>
    [JsonPropertyName("bookmarkName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BookmarkName { get; init; }
}
