using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Information about freeze panes.
/// </summary>
public record FreezePanesInfo
{
    /// <summary>
    ///     First visible row index.
    /// </summary>
    [JsonPropertyName("row")]
    public required int Row { get; init; }

    /// <summary>
    ///     First visible column index.
    /// </summary>
    [JsonPropertyName("column")]
    public required int Column { get; init; }
}
