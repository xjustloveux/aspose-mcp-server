using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Table;

/// <summary>
///     Result type for getting table content from PowerPoint presentations.
/// </summary>
public sealed record GetTableContentResult
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Gets the shape index.
    /// </summary>
    [JsonPropertyName("shapeIndex")]
    public required int ShapeIndex { get; init; }

    /// <summary>
    ///     Gets the row count.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public required int RowCount { get; init; }

    /// <summary>
    ///     Gets the column count.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }

    /// <summary>
    ///     Gets the table data as a 2D array.
    /// </summary>
    [JsonPropertyName("data")]
    public required List<List<string>> Data { get; init; }
}
