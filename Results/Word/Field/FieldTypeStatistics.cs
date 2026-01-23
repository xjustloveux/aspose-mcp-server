using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Statistics for a field type.
/// </summary>
public sealed record FieldTypeStatistics
{
    /// <summary>
    ///     Field type name.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Number of fields of this type.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }
}
