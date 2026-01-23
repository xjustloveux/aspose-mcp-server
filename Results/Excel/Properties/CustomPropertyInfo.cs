using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Information about a custom document property.
/// </summary>
public record CustomPropertyInfo
{
    /// <summary>
    ///     Name of the custom property.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Value of the custom property as a string.
    /// </summary>
    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }

    /// <summary>
    ///     Type of the custom property.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
