using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Properties;

/// <summary>
///     Built-in document properties.
/// </summary>
public sealed record BuiltInPropertiesInfo
{
    /// <summary>
    ///     Document title.
    /// </summary>
    [JsonPropertyName("title")]
    public string? Title { get; init; }

    /// <summary>
    ///     Document subject.
    /// </summary>
    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    /// <summary>
    ///     Document author.
    /// </summary>
    [JsonPropertyName("author")]
    public string? Author { get; init; }

    /// <summary>
    ///     Document keywords.
    /// </summary>
    [JsonPropertyName("keywords")]
    public string? Keywords { get; init; }

    /// <summary>
    ///     Document comments.
    /// </summary>
    [JsonPropertyName("comments")]
    public string? Comments { get; init; }

    /// <summary>
    ///     Document category.
    /// </summary>
    [JsonPropertyName("category")]
    public string? Category { get; init; }

    /// <summary>
    ///     Document company.
    /// </summary>
    [JsonPropertyName("company")]
    public string? Company { get; init; }

    /// <summary>
    ///     Document manager.
    /// </summary>
    [JsonPropertyName("manager")]
    public string? Manager { get; init; }

    /// <summary>
    ///     Document creation time in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("createdTime")]
    public required string CreatedTime { get; init; }

    /// <summary>
    ///     Document last saved time in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("lastSavedTime")]
    public required string LastSavedTime { get; init; }

    /// <summary>
    ///     Name of the last person who saved the document.
    /// </summary>
    [JsonPropertyName("lastSavedBy")]
    public string? LastSavedBy { get; init; }

    /// <summary>
    ///     Document revision number.
    /// </summary>
    [JsonPropertyName("revisionNumber")]
    public required int RevisionNumber { get; init; }
}
