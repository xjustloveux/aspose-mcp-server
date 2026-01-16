using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for getting filter status from Excel worksheets.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var filterStatusParams = ExtractFilterStatusParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, filterStatusParams.SheetIndex);
        var autoFilter = worksheet.AutoFilter;

        var rangeProperty = autoFilter.Range;
        var isFilterEnabled = !string.IsNullOrEmpty(rangeProperty) && rangeProperty.Trim() != "";

        var filterColumns = autoFilter.FilterColumns;
        var hasActiveFilters = filterColumns is { Count: > 0 };

        List<object> filterColumnsList = [];
        if (filterColumns != null)
            for (var i = 0; i < filterColumns.Count; i++)
            {
                var filterColumn = filterColumns[i];
                filterColumnsList.Add(new
                {
                    columnIndex = i,
                    filterType = filterColumn.FilterType.ToString(),
                    isDropdownVisible = filterColumn.IsDropdownVisible
                });
            }

        return JsonResult(new
        {
            worksheetName = worksheet.Name,
            isFilterEnabled,
            hasActiveFilters,
            status = isFilterEnabled
                ? hasActiveFilters
                    ? "Auto filter enabled with active criteria"
                    : "Auto filter enabled (no criteria)"
                : "Auto filter not enabled",
            filterRange = isFilterEnabled ? rangeProperty : null,
            filterColumnsCount = filterColumns?.Count ?? 0,
            filterColumns = filterColumnsList
        });
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
    private record FilterStatusParameters(int SheetIndex);
}
