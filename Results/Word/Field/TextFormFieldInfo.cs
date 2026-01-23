using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Text input form field information.
/// </summary>
public sealed record TextFormFieldInfo : FormFieldInfo
{
    /// <summary>
    ///     Text value of the form field.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }
}
