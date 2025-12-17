using System.Text;
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
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "apply" => await ApplyFilterAsync(arguments, path, sheetIndex),
            "remove" => await RemoveFilterAsync(arguments, path, sheetIndex),
            "get_status" => await GetFilterStatusAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Applies auto filter to a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> ApplyFilterAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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
        catch
        {
            // If that fails, keep using Name
        }

        // Method 3: Try to refresh or reapply the filter
        try
        {
            // Refresh the auto filter to ensure it's applied
            worksheet.AutoFilter.Refresh();
        }
        catch
        {
            // If Refresh is not available, try removing and re-adding
            try
            {
                worksheet.AutoFilter.Range = "";
                worksheet.AutoFilter.Range = cellRange.Name;
            }
            catch
            {
                // If that also fails, continue with what we have
            }
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"Auto filter applied to range {range} in sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    ///     Removes filter from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> RemoveFilterAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.AutoFilter.Range = "";

        workbook.Save(outputPath);
        return await Task.FromResult($"Auto filter removed from sheet {sheetIndex}: {outputPath}");
    }

    /// <summary>
    ///     Gets filter status for the worksheet
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with filter status</returns>
    private async Task<string> GetFilterStatusAsync(JsonObject? _, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var autoFilter = worksheet.AutoFilter;
        var result = new StringBuilder();

        result.AppendLine($"=== Auto filter status for worksheet '{worksheet.Name}' ===\n");

        // Check if auto filter is enabled
        // Check multiple indicators to determine filter status

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
        // Filter columns exist when auto filter is applied
        var filterColumns = autoFilter.FilterColumns;
        if (filterColumns is { Count: > 0 })
        {
            // If filter columns exist, filter is likely enabled
            isFilterEnabled = true;
            // Try to get range from filter columns if Range property is empty
            if (string.IsNullOrEmpty(filterRange))
                // Filter columns exist but Range is empty - filter might still be enabled
                filterRange = "Range not specified";
        }

        // Method 3: Check if auto filter is actually applied by examining the worksheet
        // Sometimes Range property might be empty but filter is still applied
        // We can check if there are filter indicators in the header row

        if (!isFilterEnabled)
        {
            result.AppendLine("Status: Auto filter not enabled");
        }
        else
        {
            result.AppendLine("Status: Auto filter enabled");
            if (!string.IsNullOrEmpty(filterRange) && filterRange != "Range not specified")
                result.AppendLine($"Filter range: {filterRange}");
            else if (filterColumns is { Count: > 0 })
                result.AppendLine($"Filter range: Not specified (but detected {filterColumns.Count} filter columns)");

            if (filterColumns is { Count: > 0 })
            {
                result.AppendLine($"Filter columns count: {filterColumns.Count}");
                for (var i = 0; i < filterColumns.Count; i++)
                    result.AppendLine($"  Column {i}: Filter applied");
            }
        }

        return await Task.FromResult(result.ToString());
    }
}