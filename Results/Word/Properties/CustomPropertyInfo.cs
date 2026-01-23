using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Properties;

/// <summary>
///     Custom document property information.
/// </summary>
public sealed record CustomPropertyInfo
{
    /// <summary>
    ///     Property value.
    /// </summary>
    [JsonPropertyName("value")]
    public object? Value { get; init; }

    /// <summary>
    ///     Property type name.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
