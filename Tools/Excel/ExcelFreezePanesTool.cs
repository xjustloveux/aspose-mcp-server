using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel freeze panes (freeze/unfreeze/get)
/// Merges: ExcelFreezePanesTool, ExcelGetFreezePanesTool
/// </summary>
public class ExcelFreezePanesTool : IAsposeTool
{
    public string Description => "Manage Excel freeze panes: freeze, unfreeze, or get freeze panes status";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'freeze', 'unfreeze', 'get'",
                @enum = new[] { "freeze", "unfreeze", "get" }
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
            row = new
            {
                type = "number",
                description = "Row index to freeze at (0-based, required for freeze)"
            },
            column = new
            {
                type = "number",
                description = "Column index to freeze at (0-based, required for freeze)"
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
            "freeze" => await FreezePanesAsync(arguments, path, sheetIndex),
            "unfreeze" => await UnfreezePanesAsync(arguments, path, sheetIndex),
            "get" => await GetFreezePanesAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> FreezePanesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var row = arguments?["row"]?.GetValue<int>() ?? throw new ArgumentException("row is required for freeze operation");
        var column = arguments?["column"]?.GetValue<int>() ?? throw new ArgumentException("column is required for freeze operation");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        // FreezePanes parameters are 1-based in Aspose.Cells
        // If user provides 0-based row/column, convert to 1-based
        worksheet.FreezePanes(row + 1, column + 1, row + 1, column + 1);
        
        // Save freeze information to custom properties for later retrieval
        // Since FirstVisibleRow/FirstVisibleColumn might not reflect freeze status,
        // we'll store it in custom properties
        var customProperties = workbook.CustomDocumentProperties;
        var freezeKey = $"FreezePanes_Sheet{sheetIndex}";
        var freezeValue = $"{row},{column}";
        
        // Remove existing freeze property if any
        try
        {
            customProperties.Remove(freezeKey);
        }
        catch
        {
            // Ignore if property doesn't exist
        }
        
        // Add new freeze property
        customProperties.Add(freezeKey, freezeValue);
        
        workbook.Save(path);
        return await Task.FromResult($"已凍結窗格 (行 {row}, 列 {column}): {path}");
    }

    private async Task<string> UnfreezePanesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.FreezePanes(0, 0, 0, 0);
        
        // Remove freeze information from custom properties
        var customProperties = workbook.CustomDocumentProperties;
        var freezeKey = $"FreezePanes_Sheet{sheetIndex}";
        try
        {
            customProperties.Remove(freezeKey);
        }
        catch
        {
            // Ignore if property doesn't exist
        }
        
        workbook.Save(path);
        return await Task.FromResult($"已取消凍結窗格: {path}");
    }

    private async Task<string> GetFreezePanesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的凍結窗格狀態 ===\n");

        // Check freeze panes status
        // Since FirstVisibleRow/FirstVisibleColumn might not reliably detect freeze status,
        // we'll use custom properties to track freeze information
        
        bool isFrozen = false;
        int frozenRow = 0;
        int frozenColumn = 0;
        
        // Method 1: Check custom properties (most reliable)
        // We store freeze information in custom properties when freeze is applied
        var customProperties = workbook.CustomDocumentProperties;
        var freezeKey = $"FreezePanes_Sheet{sheetIndex}";
        
        try
        {
            var freezeProperty = customProperties[freezeKey];
            if (freezeProperty != null)
            {
                var freezeValue = freezeProperty.Value?.ToString();
                if (!string.IsNullOrEmpty(freezeValue))
                {
                    // Parse freeze value: "row,column"
                    var parts = freezeValue.Split(',');
                    if (parts.Length == 2 && 
                        int.TryParse(parts[0], out int parsedRow) && 
                        int.TryParse(parts[1], out int parsedColumn))
                    {
                        isFrozen = true;
                        frozenRow = parsedRow;
                        frozenColumn = parsedColumn;
                    }
                }
            }
        }
        catch
        {
            // If custom property doesn't exist or can't be read, try other methods
        }
        
        // Method 2: Check FirstVisibleRow and FirstVisibleColumn as fallback
        // These might reflect the freeze position, but can be affected by scrolling
        if (!isFrozen)
        {
            var firstVisibleRow = worksheet.FirstVisibleRow;
            var firstVisibleColumn = worksheet.FirstVisibleColumn;
            
            if (firstVisibleRow > 0 || firstVisibleColumn > 0)
            {
                isFrozen = true;
                frozenRow = firstVisibleRow;
                frozenColumn = firstVisibleColumn;
            }
        }
        
        if (!isFrozen)
        {
            result.AppendLine("狀態: 未凍結窗格");
        }
        else
        {
            result.AppendLine("狀態: 已凍結窗格");
            result.AppendLine($"凍結行: {frozenRow}");
            result.AppendLine($"凍結列: {frozenColumn}");
            result.AppendLine($"凍結位置: 行 {frozenRow + 1} 和列 {frozenColumn + 1} 之前");
        }

        return await Task.FromResult(result.ToString());
    }
}
