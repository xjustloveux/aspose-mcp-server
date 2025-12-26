using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel filters (auto filter, get filter status)
///     Merges: ExcelAutoFilterTool, ExcelGetFilterStatusTool
/// </summary>
public class ExcelFilterTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel filters. Supports 3 operations: apply, remove, get_status.

Usage examples:
- Apply filter: excel_filter(operation='apply', path='book.xlsx', range='A1:C10')
- Remove filter: excel_filter(operation='remove', path='book.xlsx', range='A1:C10')
- Get filter status: excel_filter(operation='get_status', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
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
- 'apply': Apply auto filter (required params: path, range)
- 'remove': Remove auto filter (required params: path, range)
- 'get_status': Get filter status (required params: path)",
                @enum = new[] { "apply", "remove", "get_status" }
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
                description = "Cell range to apply filter (e.g., 'A1:C10', required for apply/remove)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for apply/remove operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
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
            "get_status" => await GetFilterStatusAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Applies auto filter to a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <returns>Success message</returns>
    private Task<string> ApplyFilterAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            // Set auto filter range
            // Try multiple methods to ensure filter is applied

            // Method 1: Set Range property using range name
            worksheet.AutoFilter.Range = cellRange.Name;

            // Method 2: Also try setting using the range address directly
            // Sometimes the Name property might not work, so try the address
            try
            {
                worksheet.AutoFilter.Range = range;
            }
            catch (Exception ex)
            {
                // If that fails, keep using Name
                Console.Error.WriteLine($"[WARN] Failed to set filter range directly: {ex.Message}");
            }

            // Method 3: Try to refresh or reapply the filter
            try
            {
                // Refresh the auto filter to ensure it's applied
                worksheet.AutoFilter.Refresh();
            }
            catch (Exception ex)
            {
                // If Refresh is not available, try removing and re-adding
                Console.Error.WriteLine($"[WARN] Filter refresh failed, trying alternative method: {ex.Message}");
                try
                {
                    worksheet.AutoFilter.Range = "";
                    worksheet.AutoFilter.Range = cellRange.Name;
                }
                catch (Exception ex2)
                {
                    // If that also fails, continue with what we have
                    Console.Error.WriteLine($"[WARN] Alternative filter method also failed: {ex2.Message}");
                }
            }

            workbook.Save(outputPath);

            return $"Auto filter applied to range {range} in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes filter from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> RemoveFilterAsync(string path, string outputPath, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            worksheet.AutoFilter.Range = "";

            workbook.Save(outputPath);
            return $"Auto filter removed from sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets filter status for the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with filter status</returns>
    private Task<string> GetFilterStatusAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var autoFilter = worksheet.AutoFilter;

            var isFilterEnabled = false;
            var filterRange = "";

            // Method 1: Check AutoFilter.Range property
            var rangeProperty = autoFilter.Range;
            if (!string.IsNullOrEmpty(rangeProperty) && rangeProperty.Trim() != "")
            {
                isFilterEnabled = true;
                filterRange = rangeProperty;
            }

            // Method 2: Check if there are filter columns
            var filterColumns = autoFilter.FilterColumns;
            if (filterColumns is { Count: > 0 })
            {
                isFilterEnabled = true;
                if (string.IsNullOrEmpty(filterRange))
                    filterRange = "Range not specified";
            }

            var filterColumnsList = new List<object>();
            if (filterColumns is { Count: > 0 })
                for (var i = 0; i < filterColumns.Count; i++)
                    filterColumnsList.Add(new
                    {
                        index = i,
                        status = "Filter applied"
                    });

            var result = new
            {
                worksheetName = worksheet.Name,
                isFilterEnabled,
                status = isFilterEnabled ? "Auto filter enabled" : "Auto filter not enabled",
                filterRange = isFilterEnabled ? filterRange : null,
                filterColumnsCount = filterColumns?.Count ?? 0,
                filterColumns = filterColumnsList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}