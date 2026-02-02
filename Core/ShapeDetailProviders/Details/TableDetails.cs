using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for Table elements.
/// </summary>
public sealed record TableDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the number of rows.
    /// </summary>
    [JsonPropertyName("rows")]
    public required int Rows { get; init; }

    /// <summary>
    ///     Gets the number of columns.
    /// </summary>
    [JsonPropertyName("columns")]
    public required int Columns { get; init; }

    /// <summary>
    ///     Gets the total number of cells.
    /// </summary>
    [JsonPropertyName("totalCells")]
    public required int TotalCells { get; init; }

    /// <summary>
    ///     Gets whether the first row has special formatting.
    /// </summary>
    [JsonPropertyName("firstRow")]
    public required bool FirstRow { get; init; }

    /// <summary>
    ///     Gets whether the first column has special formatting.
    /// </summary>
    [JsonPropertyName("firstCol")]
    public required bool FirstCol { get; init; }

    /// <summary>
    ///     Gets whether the last row has special formatting.
    /// </summary>
    [JsonPropertyName("lastRow")]
    public required bool LastRow { get; init; }

    /// <summary>
    ///     Gets whether the last column has special formatting.
    /// </summary>
    [JsonPropertyName("lastCol")]
    public required bool LastCol { get; init; }

    /// <summary>
    ///     Gets the table style preset name.
    /// </summary>
    [JsonPropertyName("stylePreset")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StylePreset { get; init; }

    /// <summary>
    ///     Gets the number of merged cells.
    /// </summary>
    [JsonPropertyName("mergedCellCount")]
    public required int MergedCellCount { get; init; }

    /// <summary>
    ///     Gets the merged cell information.
    /// </summary>
    [JsonPropertyName("mergedCells")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<MergedCellInfo>? MergedCells { get; init; }
}

/// <summary>
///     Information about a merged cell in a table.
/// </summary>
public sealed record MergedCellInfo
{
    /// <summary>
    ///     Gets the row index.
    /// </summary>
    [JsonPropertyName("row")]
    public required int Row { get; init; }

    /// <summary>
    ///     Gets the column index.
    /// </summary>
    [JsonPropertyName("col")]
    public required int Col { get; init; }

    /// <summary>
    ///     Gets the number of columns spanned.
    /// </summary>
    [JsonPropertyName("colSpan")]
    public required int ColSpan { get; init; }

    /// <summary>
    ///     Gets the number of rows spanned.
    /// </summary>
    [JsonPropertyName("rowSpan")]
    public required int RowSpan { get; init; }
}
