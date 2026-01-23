using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Filter;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for getting filter status from Excel worksheets.
/// </summary>
[ResultType(typeof(GetFilterStatusResult))]
public class GetExcelFilterStatusHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_status";

    /// <summary>
    ///     Gets the current filter status and details.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with filter status information.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var filterStatusParams = ExtractFilterStatusParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, filterStatusParams.SheetIndex);
        var autoFilter = worksheet.AutoFilter;

        var rangeProperty = autoFilter.Range;
        var isFilterEnabled = !string.IsNullOrEmpty(rangeProperty) && rangeProperty.Trim() != "";

        var filterColumns = autoFilter.FilterColumns;
        var hasActiveFilters = filterColumns is { Count: > 0 };

        List<FilterColumnInfo> filterColumnsList = [];
        if (filterColumns != null)
            for (var i = 0; i < filterColumns.Count; i++)
            {
                var filterColumn = filterColumns[i];
                filterColumnsList.Add(new FilterColumnInfo
                {
                    ColumnIndex = i,
                    FilterType = filterColumn.FilterType.ToString(),
                    IsDropdownVisible = filterColumn.IsDropdownVisible
                });
            }

        return new GetFilterStatusResult
        {
            WorksheetName = worksheet.Name,
            IsFilterEnabled = isFilterEnabled,
            HasActiveFilters = hasActiveFilters,
            Status = GetFilterStatusDescription(isFilterEnabled, hasActiveFilters),
            FilterRange = isFilterEnabled ? rangeProperty : null,
            FilterColumnsCount = filterColumns?.Count ?? 0,
            FilterColumns = filterColumnsList
        };
    }

    /// <summary>
    ///     Gets the filter status description string.
    /// </summary>
    /// <param name="isFilterEnabled">Whether the filter is enabled.</param>
    /// <param name="hasActiveFilters">Whether there are active filter criteria.</param>
    /// <returns>The status description string.</returns>
    private static string GetFilterStatusDescription(bool isFilterEnabled, bool hasActiveFilters)
    {
        if (!isFilterEnabled) return "Auto filter not enabled";
        if (hasActiveFilters) return "Auto filter enabled with active criteria";
        return "Auto filter enabled (no criteria)";
    }

    /// <summary>
    ///     Extracts filter status parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted filter status parameters.</returns>
    private static FilterStatusParameters ExtractFilterStatusParameters(OperationParameters parameters)
    {
        return new FilterStatusParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Parameters for get filter status operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    private sealed record FilterStatusParameters(int SheetIndex);
}
