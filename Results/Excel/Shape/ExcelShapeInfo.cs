using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Shape;

/// <summary>
///     Information about a single Excel shape.
/// </summary>
public record ExcelShapeInfo
{
    /// <summary>
    ///     Shape index in the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Shape name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Shape type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Text content of the shape.
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }

    /// <summary>
    ///     Upper-left row.
    /// </summary>
    [JsonPropertyName("upperLeftRow")]
    public int UpperLeftRow { get; init; }

    /// <summary>
    ///     Upper-left column.
    /// </summary>
    [JsonPropertyName("upperLeftColumn")]
    public int UpperLeftColumn { get; init; }

    /// <summary>
    ///     Width in pixels.
    /// </summary>
    [JsonPropertyName("width")]
    public int Width { get; init; }

    /// <summary>
    ///     Height in pixels.
    /// </summary>
    [JsonPropertyName("height")]
    public int Height { get; init; }
}
