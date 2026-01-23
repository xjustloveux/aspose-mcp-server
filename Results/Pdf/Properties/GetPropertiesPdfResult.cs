using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Properties;

/// <summary>
///     Result for getting properties from PDF documents.
/// </summary>
public record GetPropertiesPdfResult
{
    /// <summary>
    ///     Document title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Document author.
    /// </summary>
    [JsonPropertyName("author")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; init; }

    /// <summary>
    ///     Document subject.
    /// </summary>
    [JsonPropertyName("subject")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; init; }

    /// <summary>
    ///     Document keywords.
    /// </summary>
    [JsonPropertyName("keywords")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Keywords { get; init; }

    /// <summary>
    ///     Document creator application.
    /// </summary>
    [JsonPropertyName("creator")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Creator { get; init; }

    /// <summary>
    ///     Document producer application.
    /// </summary>
    [JsonPropertyName("producer")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Producer { get; init; }

    /// <summary>
    ///     Creation date.
    /// </summary>
    [JsonPropertyName("creationDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CreationDate { get; init; }

    /// <summary>
    ///     Modification date.
    /// </summary>
    [JsonPropertyName("modificationDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ModificationDate { get; init; }

    /// <summary>
    ///     Total number of pages.
    /// </summary>
    [JsonPropertyName("totalPages")]
    public required int TotalPages { get; init; }

    /// <summary>
    ///     Whether document is encrypted.
    /// </summary>
    [JsonPropertyName("isEncrypted")]
    public required bool IsEncrypted { get; init; }

    /// <summary>
    ///     Whether document is linearized.
    /// </summary>
    [JsonPropertyName("isLinearized")]
    public required bool IsLinearized { get; init; }
}
