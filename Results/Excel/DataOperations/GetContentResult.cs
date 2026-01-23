using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Result for getting content from Excel worksheets.
/// </summary>
public record GetContentResult
{
    /// <summary>
    ///     The content rows as a list of dictionaries.
    ///     Each dictionary represents a row with column names as keys.
    /// </summary>
    [JsonPropertyName("rows")]
    public required IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows { get; init; }
}
