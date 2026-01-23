using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.CellAddress;

/// <summary>
///     Result for cell address conversion operations.
/// </summary>
public record CellAddressResult
{
    /// <summary>
    ///     Cell address in A1 notation (e.g., "A1", "B2", "AA100").
    /// </summary>
    [JsonPropertyName("a1Notation")]
    public required string A1Notation { get; init; }

    /// <summary>
    ///     Zero-based row index.
    /// </summary>
    [JsonPropertyName("row")]
    public required int Row { get; init; }

    /// <summary>
    ///     Zero-based column index.
    /// </summary>
    [JsonPropertyName("column")]
    public required int Column { get; init; }
}
