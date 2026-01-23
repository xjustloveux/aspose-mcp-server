using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Result type for getting statistics from PowerPoint presentations.
/// </summary>
public sealed record GetStatisticsResult
{
    /// <summary>
    ///     Gets the total number of slides.
    /// </summary>
    [JsonPropertyName("totalSlides")]
    public required int TotalSlides { get; init; }

    /// <summary>
    ///     Gets the total number of hidden slides.
    /// </summary>
    [JsonPropertyName("totalHiddenSlides")]
    public required int TotalHiddenSlides { get; init; }

    /// <summary>
    ///     Gets the total number of layouts.
    /// </summary>
    [JsonPropertyName("totalLayouts")]
    public required int TotalLayouts { get; init; }

    /// <summary>
    ///     Gets the total number of masters.
    /// </summary>
    [JsonPropertyName("totalMasters")]
    public required int TotalMasters { get; init; }

    /// <summary>
    ///     Gets the slide size information.
    /// </summary>
    [JsonPropertyName("slideSize")]
    public required GetStatisticsSizeInfo SlideSize { get; init; }

    /// <summary>
    ///     Gets the total number of shapes.
    /// </summary>
    [JsonPropertyName("totalShapes")]
    public required int TotalShapes { get; init; }

    /// <summary>
    ///     Gets the total number of text characters.
    /// </summary>
    [JsonPropertyName("totalTextCharacters")]
    public required int TotalTextCharacters { get; init; }

    /// <summary>
    ///     Gets the total number of images.
    /// </summary>
    [JsonPropertyName("totalImages")]
    public required int TotalImages { get; init; }

    /// <summary>
    ///     Gets the total number of tables.
    /// </summary>
    [JsonPropertyName("totalTables")]
    public required int TotalTables { get; init; }

    /// <summary>
    ///     Gets the total number of charts.
    /// </summary>
    [JsonPropertyName("totalCharts")]
    public required int TotalCharts { get; init; }

    /// <summary>
    ///     Gets the total number of SmartArt objects.
    /// </summary>
    [JsonPropertyName("totalSmartArt")]
    public required int TotalSmartArt { get; init; }

    /// <summary>
    ///     Gets the total number of audio objects.
    /// </summary>
    [JsonPropertyName("totalAudio")]
    public required int TotalAudio { get; init; }

    /// <summary>
    ///     Gets the total number of video objects.
    /// </summary>
    [JsonPropertyName("totalVideo")]
    public required int TotalVideo { get; init; }

    /// <summary>
    ///     Gets the total number of animations.
    /// </summary>
    [JsonPropertyName("totalAnimations")]
    public required int TotalAnimations { get; init; }

    /// <summary>
    ///     Gets the total number of hyperlinks.
    /// </summary>
    [JsonPropertyName("totalHyperlinks")]
    public required int TotalHyperlinks { get; init; }
}
