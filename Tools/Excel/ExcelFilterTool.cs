using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel filters (auto filter, custom filter, get filter status).
///     Merges: ExcelAutoFilterTool, ExcelGetFilterStatusTool.
/// </summary>
[McpServerToolType]
public class ExcelFilterTool
{
    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelFilterTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelFilterTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel filter operation (apply, remove, filter, or get_status).
    /// </summary>
    /// <param name="operation">The operation to perform: apply, remove, filter, or get_status.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range to apply filter (e.g., 'A1:C10', required for apply/filter).</param>
    /// <param name="columnIndex">Column index within filter range to apply criteria (0-based, required for filter).</param>
    /// <param name="criteria">Filter criteria value (required for filter operation).</param>
    /// <param name="filterOperator">Filter operator for custom filter.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_status operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_filter")]
    [Description(@"Manage Excel filters. Supports 4 operations: apply, remove, filter, get_status.

Usage examples:
- Apply auto filter: excel_filter(operation='apply', path='book.xlsx', range='A1:C10')
- Remove filter: excel_filter(operation='remove', path='book.xlsx')
- Filter by value: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=0, criteria='Completed')
- Filter by custom: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=1, filterOperator='GreaterThan', criteria='100')
- Get filter status: excel_filter(operation='get_status', path='book.xlsx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'apply': Apply auto filter dropdown buttons (required params: path, range)
- 'remove': Remove auto filter completely (required params: path)
- 'filter': Apply filter criteria to a column (required params: path, range, columnIndex, criteria)
- 'get_status': Get filter status with details (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range to apply filter (e.g., 'A1:C10', required for apply/filter)")]
        string? range = null,
        [Description("Column index within filter range to apply criteria (0-based, required for filter)")]
        int columnIndex = 0,
        [Description("Filter criteria value (required for filter operation)")]
        string? criteria = null,
        [Description("Filter operator for custom filter (optional, default: 'Equal')")]
        string filterOperator = "Equal")
    {
        return operation.ToLower() switch
        {
            "apply" => ApplyFilter(path, sessionId, outputPath, sheetIndex, range),
            "remove" => RemoveFilter(path, sessionId, outputPath, sheetIndex),
            "filter" => FilterByValue(path, sessionId, outputPath, sheetIndex, range, columnIndex, criteria,
                filterOperator),
            "get_status" => GetFilterStatus(path, sessionId, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Applies auto filter dropdown buttons to a range.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to apply filter to.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string ApplyFilter(string? path, string? sessionId, string? outputPath, int sheetIndex, string? range)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for apply operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, range);
        worksheet.AutoFilter.Range = range;

        ctx.Save(outputPath);
        return $"Auto filter applied to range {range} in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Removes auto filter from the worksheet.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string RemoveFilter(string? path, string? sessionId, string? outputPath, int sheetIndex)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        worksheet.RemoveAutoFilter();

        ctx.Save(outputPath);
        return $"Auto filter removed from sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Applies filter criteria to a specific column.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range for the filter.</param>
    /// <param name="columnIndex">The column index to apply criteria to.</param>
    /// <param name="criteria">The filter criteria value.</param>
    /// <param name="filterOperatorStr">The filter operator type.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string FilterByValue(string? path, string? sessionId, string? outputPath, int sheetIndex, string? range,
        int columnIndex, string? criteria, string filterOperatorStr)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for filter operation");
        if (string.IsNullOrEmpty(criteria))
            throw new ArgumentException("criteria is required for filter operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, range);
        worksheet.AutoFilter.Range = range;

        var filterOperator = ParseFilterOperator(filterOperatorStr);

        if (filterOperator == FilterOperatorType.Equal)
        {
            worksheet.AutoFilter.Filter(columnIndex, criteria);
        }
        else
        {
            object criteriaValue = criteria;
            if (IsNumericOperator(filterOperator) && double.TryParse(criteria, out var numericValue))
                criteriaValue = numericValue;

            worksheet.AutoFilter.Custom(columnIndex, filterOperator, criteriaValue);
        }

        worksheet.AutoFilter.Refresh();

        ctx.Save(outputPath);
        return
            $"Filter applied to column {columnIndex} with criteria '{criteria}' (operator: {filterOperatorStr}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets the current filter status and details.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the filter status information.</returns>
    private string GetFilterStatus(string? path, string? sessionId, int sheetIndex)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);
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

        var result = new
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
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Determines if the filter operator is a numeric comparison operator.
    /// </summary>
    /// <param name="op">The filter operator type to check.</param>
    /// <returns>True if the operator requires numeric comparison (GreaterThan, LessThan, etc.); otherwise false.</returns>
    private static bool IsNumericOperator(FilterOperatorType op)
    {
        return op is FilterOperatorType.GreaterThan or FilterOperatorType.GreaterOrEqual
            or FilterOperatorType.LessThan or FilterOperatorType.LessOrEqual;
    }

    /// <summary>
    ///     Parses filter operator string to FilterOperatorType enum.
    /// </summary>
    /// <param name="operatorStr">The filter operator string.</param>
    /// <returns>The corresponding FilterOperatorType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the operator string is not supported.</exception>
    private static FilterOperatorType ParseFilterOperator(string operatorStr)
    {
        return operatorStr switch
        {
            "Equal" => FilterOperatorType.Equal,
            "NotEqual" => FilterOperatorType.NotEqual,
            "GreaterThan" => FilterOperatorType.GreaterThan,
            "GreaterOrEqual" => FilterOperatorType.GreaterOrEqual,
            "LessThan" => FilterOperatorType.LessThan,
            "LessOrEqual" => FilterOperatorType.LessOrEqual,
            "Contains" => FilterOperatorType.Contains,
            "NotContains" => FilterOperatorType.NotContains,
            "BeginsWith" => FilterOperatorType.BeginsWith,
            "EndsWith" => FilterOperatorType.EndsWith,
            _ => throw new ArgumentException($"Unsupported filter operator: {operatorStr}")
        };
    }
}