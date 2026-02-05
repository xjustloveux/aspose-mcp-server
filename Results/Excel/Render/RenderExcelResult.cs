using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Render;

/// <summary>
///     Result for Excel render operations.
/// </summary>
public record RenderExcelResult
{
    /// <summary>
    ///     Output file path(s).
    /// </summary>
    [JsonPropertyName("outputPaths")]
    public required IReadOnlyList<string> OutputPaths { get; init; }

    /// <summary>
    ///     Number of pages/images rendered.
    /// </summary>
    [JsonPropertyName("pageCount")]
    public required int PageCount { get; init; }

    /// <summary>
    ///     Image format used.
    /// </summary>
    [JsonPropertyName("format")]
    public required string Format { get; init; }

    /// <summary>
    ///     Result message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
