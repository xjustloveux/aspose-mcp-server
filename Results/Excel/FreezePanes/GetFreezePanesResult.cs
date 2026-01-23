using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.FreezePanes;

/// <summary>
///     Result for getting freeze panes status from Excel workbooks.
/// </summary>
public record GetFreezePanesResult
{
    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     Whether panes are frozen.
    /// </summary>
    [JsonPropertyName("isFrozen")]
    public required bool IsFrozen { get; init; }

    /// <summary>
    ///     Frozen row index.
    /// </summary>
    [JsonPropertyName("frozenRow")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FrozenRow { get; init; }

    /// <summary>
    ///     Frozen column index.
    /// </summary>
    [JsonPropertyName("frozenColumn")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FrozenColumn { get; init; }

    /// <summary>
    ///     Number of frozen rows.
    /// </summary>
    [JsonPropertyName("frozenRows")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FrozenRows { get; init; }

    /// <summary>
    ///     Number of frozen columns.
    /// </summary>
    [JsonPropertyName("frozenColumns")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FrozenColumns { get; init; }

    /// <summary>
    ///     Status message.
    /// </summary>
    [JsonPropertyName("status")]
    public required string Status { get; init; }
}
