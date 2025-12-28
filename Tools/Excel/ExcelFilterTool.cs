using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel filters (auto filter, custom filter, get filter status).
///     Merges: ExcelAutoFilterTool, ExcelGetFilterStatusTool.
/// </summary>
public class ExcelFilterTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel filters. Supports 4 operations: apply, remove, filter, get_status.

Usage examples:
- Apply auto filter: excel_filter(operation='apply', path='book.xlsx', range='A1:C10')
- Remove filter: excel_filter(operation='remove', path='book.xlsx')
- Filter by value: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=0, criteria='Completed')
- Filter by custom: excel_filter(operation='filter', path='book.xlsx', range='A1:C10', columnIndex=1, filterOperator='GreaterThan', criteria='100')
- Get filter status: excel_filter(operation='get_status', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'apply': Apply auto filter dropdown buttons (required params: path, range)
- 'remove': Remove auto filter completely (required params: path)
- 'filter': Apply filter criteria to a column (required params: path, range, columnIndex, criteria)
- 'get_status': Get filter status with details (required params: path)",
                @enum = new[] { "apply", "remove", "filter", "get_status" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to apply filter (e.g., 'A1:C10', required for apply/filter)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index within filter range to apply criteria (0-based, required for filter)"
            },
            criteria = new
            {
                type = "string",
                description = "Filter criteria value (required for filter operation)"
            },
            filterOperator = new
            {
                type = "string",
                description =
                    "Filter operator for custom filter (optional, default: 'Equal'). Use with numeric/date criteria.",
                @enum = new[]
                {
                    "Equal", "NotEqual", "GreaterThan", "GreaterOrEqual", "LessThan", "LessOrEqual", "Contains",
                    "NotContains", "BeginsWith", "EndsWith"
                }
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for apply/remove/filter operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "apply" => await ApplyFilterAsync(path, outputPath, sheetIndex, arguments),
            "remove" => await RemoveFilterAsync(path, outputPath, sheetIndex),
            "filter" => await FilterByValueAsync(path, outputPath, sheetIndex, arguments),
            "get_status" => await GetFilterStatusAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Applies auto filter dropdown buttons to a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range.</param>
    /// <returns>Success message.</returns>
    private Task<string> ApplyFilterAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            ExcelHelper.CreateRange(worksheet.Cells, range);

            worksheet.AutoFilter.Range = range;

            workbook.Save(outputPath);

            return $"Auto filter applied to range {range} in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes auto filter completely from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>Success message.</returns>
    private Task<string> RemoveFilterAsync(string path, string outputPath, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            worksheet.RemoveAutoFilter();

            workbook.Save(outputPath);
            return $"Auto filter removed from sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies filter criteria to a specific column.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range, columnIndex, criteria, optional filterOperator.</param>
    /// <returns>Success message with filter details.</returns>
    private Task<string> FilterByValueAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var criteria = ArgumentHelper.GetString(arguments, "criteria");
            var filterOperatorStr = ArgumentHelper.GetStringNullable(arguments, "filterOperator") ?? "Equal";

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            ExcelHelper.CreateRange(worksheet.Cells, range);

            worksheet.AutoFilter.Range = range;

            var filterOperator = ParseFilterOperator(filterOperatorStr);

            if (filterOperator == FilterOperatorType.Equal)
                worksheet.AutoFilter.Filter(columnIndex, criteria);
            else
                worksheet.AutoFilter.Custom(columnIndex, filterOperator, criteria);

            workbook.Save(outputPath);

            return
                $"Filter applied to column {columnIndex} with criteria '{criteria}' (operator: {filterOperatorStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets detailed filter status for the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with detailed filter status.</returns>
    private Task<string> GetFilterStatusAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var autoFilter = worksheet.AutoFilter;

            var rangeProperty = autoFilter.Range;
            var isFilterEnabled = !string.IsNullOrEmpty(rangeProperty) && rangeProperty.Trim() != "";

            var filterColumns = autoFilter.FilterColumns;
            var hasActiveFilters = filterColumns is { Count: > 0 };

            var filterColumnsList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Parses filter operator string to FilterOperatorType enum.
    /// </summary>
    /// <param name="operatorStr">Operator string.</param>
    /// <returns>FilterOperatorType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown if operator is not supported.</exception>
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