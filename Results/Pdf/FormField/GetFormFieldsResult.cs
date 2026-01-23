using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.FormField;

/// <summary>
///     Result for getting form fields from PDF documents.
/// </summary>
public record GetFormFieldsResult
{
    /// <summary>
    ///     Number of form fields returned.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Total number of form fields in document.
    /// </summary>
    [JsonPropertyName("totalCount")]
    public required int TotalCount { get; init; }

    /// <summary>
    ///     Whether the result is truncated.
    /// </summary>
    [JsonPropertyName("truncated")]
    public required bool Truncated { get; init; }

    /// <summary>
    ///     List of form field information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<PdfFormFieldInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no form fields found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
