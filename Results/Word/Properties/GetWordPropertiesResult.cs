using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Properties;

/// <summary>
///     Result for getting Word document properties.
/// </summary>
public sealed record GetWordPropertiesResult
{
    /// <summary>
    ///     Built-in document properties.
    /// </summary>
    [JsonPropertyName("builtInProperties")]
    public required BuiltInPropertiesInfo BuiltInProperties { get; init; }

    /// <summary>
    ///     Document statistics.
    /// </summary>
    [JsonPropertyName("statistics")]
    public required StatisticsInfo Statistics { get; init; }

    /// <summary>
    ///     Custom document properties (null if none exist).
    /// </summary>
    [JsonPropertyName("customProperties")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyDictionary<string, CustomPropertyInfo>? CustomProperties { get; init; }
}
