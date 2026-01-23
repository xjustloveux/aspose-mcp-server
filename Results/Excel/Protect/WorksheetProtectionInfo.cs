using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Protect;

/// <summary>
///     Protection information for a single worksheet.
/// </summary>
public record WorksheetProtectionInfo
{
    /// <summary>
    ///     Zero-based index of the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Whether the worksheet is protected.
    /// </summary>
    [JsonPropertyName("isProtected")]
    public required bool IsProtected { get; init; }

    /// <summary>
    ///     Allow selecting locked cell.
    /// </summary>
    [JsonPropertyName("allowSelectingLockedCell")]
    public required bool AllowSelectingLockedCell { get; init; }

    /// <summary>
    ///     Allow selecting unlocked cell.
    /// </summary>
    [JsonPropertyName("allowSelectingUnlockedCell")]
    public required bool AllowSelectingUnlockedCell { get; init; }

    /// <summary>
    ///     Allow formatting cell.
    /// </summary>
    [JsonPropertyName("allowFormattingCell")]
    public required bool AllowFormattingCell { get; init; }

    /// <summary>
    ///     Allow formatting column.
    /// </summary>
    [JsonPropertyName("allowFormattingColumn")]
    public required bool AllowFormattingColumn { get; init; }

    /// <summary>
    ///     Allow formatting row.
    /// </summary>
    [JsonPropertyName("allowFormattingRow")]
    public required bool AllowFormattingRow { get; init; }

    /// <summary>
    ///     Allow inserting column.
    /// </summary>
    [JsonPropertyName("allowInsertingColumn")]
    public required bool AllowInsertingColumn { get; init; }

    /// <summary>
    ///     Allow inserting row.
    /// </summary>
    [JsonPropertyName("allowInsertingRow")]
    public required bool AllowInsertingRow { get; init; }

    /// <summary>
    ///     Allow inserting hyperlink.
    /// </summary>
    [JsonPropertyName("allowInsertingHyperlink")]
    public required bool AllowInsertingHyperlink { get; init; }

    /// <summary>
    ///     Allow deleting column.
    /// </summary>
    [JsonPropertyName("allowDeletingColumn")]
    public required bool AllowDeletingColumn { get; init; }

    /// <summary>
    ///     Allow deleting row.
    /// </summary>
    [JsonPropertyName("allowDeletingRow")]
    public required bool AllowDeletingRow { get; init; }

    /// <summary>
    ///     Allow sorting.
    /// </summary>
    [JsonPropertyName("allowSorting")]
    public required bool AllowSorting { get; init; }

    /// <summary>
    ///     Allow filtering.
    /// </summary>
    [JsonPropertyName("allowFiltering")]
    public required bool AllowFiltering { get; init; }

    /// <summary>
    ///     Allow using pivot table.
    /// </summary>
    [JsonPropertyName("allowUsingPivotTable")]
    public required bool AllowUsingPivotTable { get; init; }

    /// <summary>
    ///     Allow editing object.
    /// </summary>
    [JsonPropertyName("allowEditingObject")]
    public required bool AllowEditingObject { get; init; }

    /// <summary>
    ///     Allow editing scenario.
    /// </summary>
    [JsonPropertyName("allowEditingScenario")]
    public required bool AllowEditingScenario { get; init; }
}
