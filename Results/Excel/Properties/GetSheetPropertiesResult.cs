using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Result for getting worksheet properties from Excel files.
/// </summary>
public record GetSheetPropertiesResult
{
    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Zero-based index of the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Whether the worksheet is visible.
    /// </summary>
    [JsonPropertyName("isVisible")]
    public required bool IsVisible { get; init; }

    /// <summary>
    ///     Tab color of the worksheet.
    /// </summary>
    [JsonPropertyName("tabColor")]
    public required string TabColor { get; init; }

    /// <summary>
    ///     Whether the worksheet is currently selected.
    /// </summary>
    [JsonPropertyName("isSelected")]
    public required bool IsSelected { get; init; }

    /// <summary>
    ///     Number of rows containing data.
    /// </summary>
    [JsonPropertyName("dataRowCount")]
    public required int DataRowCount { get; init; }

    /// <summary>
    ///     Number of columns containing data.
    /// </summary>
    [JsonPropertyName("dataColumnCount")]
    public required int DataColumnCount { get; init; }

    /// <summary>
    ///     Whether the worksheet is password protected.
    /// </summary>
    [JsonPropertyName("isProtected")]
    public required bool IsProtected { get; init; }

    /// <summary>
    ///     Number of comments in the worksheet.
    /// </summary>
    [JsonPropertyName("commentsCount")]
    public required int CommentsCount { get; init; }

    /// <summary>
    ///     Number of charts in the worksheet.
    /// </summary>
    [JsonPropertyName("chartsCount")]
    public required int ChartsCount { get; init; }

    /// <summary>
    ///     Number of pictures in the worksheet.
    /// </summary>
    [JsonPropertyName("picturesCount")]
    public required int PicturesCount { get; init; }

    /// <summary>
    ///     Number of hyperlinks in the worksheet.
    /// </summary>
    [JsonPropertyName("hyperlinksCount")]
    public required int HyperlinksCount { get; init; }

    /// <summary>
    ///     Print settings for the worksheet.
    /// </summary>
    [JsonPropertyName("printSettings")]
    public required PrintSettingsInfo PrintSettings { get; init; }
}
