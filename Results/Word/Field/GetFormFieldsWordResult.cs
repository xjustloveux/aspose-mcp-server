using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Field;

/// <summary>
///     Result for getting form fields from Word documents.
/// </summary>
public sealed record GetFormFieldsWordResult
{
    /// <summary>
    ///     Total number of form fields.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of form fields.
    /// </summary>
    [JsonPropertyName("formFields")]
    public required IReadOnlyList<FormFieldInfo> FormFields { get; init; }
}
