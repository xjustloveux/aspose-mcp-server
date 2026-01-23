using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Result for getting fields from Word documents.
/// </summary>
public sealed record GetFieldsWordResult
{
    /// <summary>
    ///     Total number of fields.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of fields.
    /// </summary>
    [JsonPropertyName("fields")]
    public required IReadOnlyList<FieldInfo> Fields { get; init; }

    /// <summary>
    ///     Statistics by field type.
    /// </summary>
    [JsonPropertyName("statisticsByType")]
    public required IReadOnlyList<FieldTypeStatistics> StatisticsByType { get; init; }
}
