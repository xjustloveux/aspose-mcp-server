using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.FormField;

/// <summary>
///     Information about a single form field.
/// </summary>
public record PdfFormFieldInfo
{
    /// <summary>
    ///     Field name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Field type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Field value for text fields.
    /// </summary>
    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }

    /// <summary>
    ///     Checked state for checkboxes.
    /// </summary>
    [JsonPropertyName("checked")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Checked { get; init; }

    /// <summary>
    ///     Selected index for radio buttons.
    /// </summary>
    [JsonPropertyName("selected")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Selected { get; init; }
}
