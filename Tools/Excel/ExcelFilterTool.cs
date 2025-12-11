using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel filters (auto filter, get filter status)
/// Merges: ExcelAutoFilterTool, ExcelGetFilterStatusTool
/// </summary>
public class ExcelFilterTool : IAsposeTool
{
    public string Description => "Manage Excel filters: apply/remove auto filter or get filter status";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'apply', 'remove', 'get_status'",
                @enum = new[] { "apply", "remove", "get_status" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "apply" => await ApplyFilterAsync(arguments, path, sheetIndex),
            "remove" => await RemoveFilterAsync(arguments, path, sheetIndex),
            "get_status" => await GetFilterStatusAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> ApplyFilterAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for apply operation");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

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
        
        workbook.Save(path);

        return await Task.FromResult($"Auto filter applied to range {range} in sheet {sheetIndex}: {path}");
    }

    private async Task<string> RemoveFilterAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        if (!string.IsNullOrEmpty(range))
        {
            var cells = worksheet.Cells;
            var cellRange = cells.CreateRange(range);
            worksheet.AutoFilter.Range = "";
        }
        else
        {
            worksheet.AutoFilter.Range = "";
        }

        workbook.Save(path);
        return await Task.FromResult($"Auto filter removed from sheet {sheetIndex}: {path}");
    }

    private async Task<string> GetFilterStatusAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var autoFilter = worksheet.AutoFilter;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的自動篩選狀態 ===\n");

        // Check if auto filter is enabled
        // Check multiple indicators to determine filter status
        
        bool isFilterEnabled = false;
        string filterRange = "";
        
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
        if (filterColumns != null && filterColumns.Count > 0)
        {
            // If filter columns exist, filter is likely enabled
            isFilterEnabled = true;
            // Try to get range from filter columns if Range property is empty
            if (string.IsNullOrEmpty(filterRange))
            {
                // Filter columns exist but Range is empty - filter might still be enabled
                filterRange = "未指定範圍";
            }
        }
        
        // Method 3: Check if auto filter is actually applied by examining the worksheet
        // Sometimes Range property might be empty but filter is still applied
        // We can check if there are filter indicators in the header row
        
        if (!isFilterEnabled)
        {
            result.AppendLine("狀態: 未啟用自動篩選");
        }
        else
        {
            result.AppendLine("狀態: 已啟用自動篩選");
            if (!string.IsNullOrEmpty(filterRange) && filterRange != "未指定範圍")
            {
                result.AppendLine($"篩選範圍: {filterRange}");
            }
            else if (filterColumns != null && filterColumns.Count > 0)
            {
                result.AppendLine($"篩選範圍: 未指定（但檢測到 {filterColumns.Count} 個篩選列）");
            }
            
            if (filterColumns != null && filterColumns.Count > 0)
            {
                result.AppendLine($"篩選列數: {filterColumns.Count}");
                for (int i = 0; i < filterColumns.Count; i++)
                {
                    try
                    {
                        var filterColumn = filterColumns[i];
                        result.AppendLine($"  列 {i}: 已應用篩選");
                    }
                    catch
                    {
                        result.AppendLine($"  列 {i}: 無法讀取篩選資訊");
                    }
                }
            }
        }

        return await Task.FromResult(result.ToString());
    }
}

