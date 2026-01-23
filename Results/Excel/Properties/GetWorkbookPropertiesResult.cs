using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Result for getting workbook properties from Excel files.
/// </summary>
public record GetWorkbookPropertiesResult
{
    /// <summary>
    ///     Title of the workbook.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Subject of the workbook.
    /// </summary>
    [JsonPropertyName("subject")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; init; }

    /// <summary>
    ///     Author of the workbook.
    /// </summary>
    [JsonPropertyName("author")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; init; }

    /// <summary>
    ///     Keywords associated with the workbook.
    /// </summary>
    [JsonPropertyName("keywords")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Keywords { get; init; }

    /// <summary>
    ///     Comments about the workbook.
    /// </summary>
    [JsonPropertyName("comments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Comments { get; init; }

    /// <summary>
    ///     Category of the workbook.
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
    ///     Creation date and time in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("created")]
    public required string Created { get; init; }

    /// <summary>
    ///     Last modification date and time in ISO 8601 format.
    /// </summary>
    [JsonPropertyName("modified")]
    public required string Modified { get; init; }

    /// <summary>
    ///     Name of the user who last saved the workbook.
    /// </summary>
    [JsonPropertyName("lastSavedBy")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LastSavedBy { get; init; }

    /// <summary>
    ///     Revision number.
    /// </summary>
    [JsonPropertyName("revision")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Revision { get; init; }

    /// <summary>
    ///     Total number of worksheets in the workbook.
    /// </summary>
    [JsonPropertyName("totalSheets")]
    public required int TotalSheets { get; init; }

    /// <summary>
    ///     List of custom document properties.
    /// </summary>
    [JsonPropertyName("customProperties")]
    public required IReadOnlyList<CustomPropertyInfo> CustomProperties { get; init; }
}
