using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Properties;

/// <summary>
///     Result for getting properties from PowerPoint presentations.
/// </summary>
public record GetPropertiesPptResult
{
    /// <summary>
    ///     Document title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Document subject.
    /// </summary>
    [JsonPropertyName("subject")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; init; }

    /// <summary>
    ///     Document author.
    /// </summary>
    [JsonPropertyName("author")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; init; }

    /// <summary>
    ///     Document keywords.
    /// </summary>
    [JsonPropertyName("keywords")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Keywords { get; init; }

    /// <summary>
    ///     Document comments.
    /// </summary>
    [JsonPropertyName("comments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comments { get; init; }

    /// <summary>
    ///     Document category.
    /// </summary>
    [JsonPropertyName("category")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Category { get; init; }

    /// <summary>
    ///     Company name.
    /// </summary>
    [JsonPropertyName("company")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Company { get; init; }

    /// <summary>
    ///     Manager name.
    /// </summary>
    [JsonPropertyName("manager")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Manager { get; init; }

    /// <summary>
    ///     Created time.
    /// </summary>
    [JsonPropertyName("createdTime")]
    public DateTime CreatedTime { get; init; }

    /// <summary>
    ///     Last saved time.
    /// </summary>
    [JsonPropertyName("lastSavedTime")]
    public DateTime LastSavedTime { get; init; }

    /// <summary>
    ///     Revision number.
    /// </summary>
    [JsonPropertyName("revisionNumber")]
    public required int RevisionNumber { get; init; }
}
