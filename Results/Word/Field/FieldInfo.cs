using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Information about a single field.
/// </summary>
public sealed record FieldInfo
{
    /// <summary>
    ///     Zero-based index of the field.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Field type (Date, PageRef, Hyperlink, etc.).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Field code (when includeCode is true).
    /// </summary>
    [JsonPropertyName("code")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Code { get; init; }

    /// <summary>
    ///     Field result/display value (when includeResult is true).
    /// </summary>
    [JsonPropertyName("result")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
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
    ///     Extra information for specific field types (hyperlink address, bookmark name, etc.).
    /// </summary>
    [JsonPropertyName("extraInfo")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ExtraInfo { get; init; }
}
