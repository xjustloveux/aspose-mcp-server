using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel hyperlinks (add, edit, delete, get)
/// Merges: ExcelAddHyperlinkTool, ExcelEditHyperlinkTool, ExcelDeleteHyperlinkTool, ExcelGetHyperlinksTool
/// </summary>
public class ExcelHyperlinkTool : IAsposeTool
{
    public string Description => @"Manage Excel hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: excel_hyperlink(operation='add', path='book.xlsx', cell='A1', url='https://example.com', displayText='Link')
- Edit hyperlink: excel_hyperlink(operation='edit', path='book.xlsx', cell='A1', url='https://newurl.com')
- Delete hyperlink: excel_hyperlink(operation='delete', path='book.xlsx', cell='A1')
- Get hyperlinks: excel_hyperlink(operation='get', path='book.xlsx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a hyperlink (required params: path, cell, url)
- 'edit': Edit a hyperlink (required params: path, cell, url)
- 'delete': Delete a hyperlink (required params: path, cell)
- 'get': Get all hyperlinks (required params: path)",
                @enum = new[] { "add", "edit", "delete", "get" }
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
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1', required for add, optional for edit/delete as alternative to hyperlinkIndex)"
            },
            url = new
            {
                type = "string",
                description = "URL or file path for the hyperlink (required for add)"
            },
            displayText = new
            {
                type = "string",
                description = "Display text for the hyperlink (optional)"
            },
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index (0-based, required for edit/delete, or use cell as alternative)"
            },
            address = new
            {
                type = "string",
                description = "New hyperlink address (optional, for edit)"
            },
            textToDisplay = new
            {
                type = "string",
                description = "New display text (optional, for edit)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(arguments, path, sheetIndex),
            "edit" => await EditHyperlinkAsync(arguments, path, sheetIndex),
            "delete" => await DeleteHyperlinkAsync(arguments, path, sheetIndex),
            "get" => await GetHyperlinksAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a hyperlink to a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell, address, and optional screenTip, textToDisplay</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with hyperlink details</returns>
    private async Task<string> AddHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var url = ArgumentHelper.GetString(arguments, "url", "url");
        var displayText = arguments?["displayText"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cellObj = worksheet.Cells[cell];

        if (!string.IsNullOrEmpty(displayText))
        {
            cellObj.PutValue(displayText);
        }

        worksheet.Hyperlinks.Add(cell, 1, 1, url);
        workbook.Save(path);

        return await Task.FromResult($"單元格 {cell} 已添加超連結: {url}");
    }

    /// <summary>
    /// Edits an existing hyperlink
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and optional address, screenTip, textToDisplay</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with updated hyperlink details</returns>
    private async Task<string> EditHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int?>();
        var cell = arguments?["cell"]?.GetValue<string>();
        var address = arguments?["address"]?.GetValue<string>();
        var textToDisplay = arguments?["textToDisplay"]?.GetValue<string>();

        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
        {
            throw new ArgumentException("Either hyperlinkIndex or cell is required for edit operation");
        }

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;
        
        int? foundIndex = hyperlinkIndex;
        
        // If cell is provided, find the hyperlink index by cell address
        if (!hyperlinkIndex.HasValue && !string.IsNullOrEmpty(cell))
        {
            foundIndex = null;
            int rowIndex, colIndex;
            try
            {
                CellsHelper.CellNameToIndex(cell, out rowIndex, out colIndex);
            }
            catch
            {
                throw new ArgumentException($"Invalid cell address: {cell}");
            }
            
            for (int i = 0; i < hyperlinks.Count; i++)
            {
                var link = hyperlinks[i];
                var area = link.Area;
                if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                    colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                {
                    foundIndex = i;
                    break;
                }
            }
            
            if (!foundIndex.HasValue)
            {
                throw new ArgumentException($"No hyperlink found at cell {cell}");
            }
        }
        
        if (!foundIndex.HasValue)
        {
            throw new ArgumentException("hyperlinkIndex is required");
        }
        
        int index = foundIndex.Value;
        if (index < 0 || index >= hyperlinks.Count)
        {
            throw new ArgumentException($"超連結索引 {index} 超出範圍 (工作表共有 {hyperlinks.Count} 個超連結)");
        }

        var hyperlink = hyperlinks[index];
        var oldAddress = hyperlink.Address ?? "";
        var oldText = hyperlink.TextToDisplay ?? "";
        
        if (!string.IsNullOrEmpty(address))
        {
            hyperlink.Address = address;
        }
        
        if (!string.IsNullOrEmpty(textToDisplay))
        {
            hyperlink.TextToDisplay = textToDisplay;
        }
        
        workbook.Save(path);
        
        var result = $"成功編輯超連結 #{index}";
        if (!string.IsNullOrEmpty(cell))
        {
            result += $" (單元格: {cell})";
        }
        result += "\n";
        result += $"舊地址: {oldAddress}\n";
        result += $"新地址: {hyperlink.Address ?? oldAddress}\n";
        result += $"舊顯示文字: {oldText}\n";
        result += $"新顯示文字: {hyperlink.TextToDisplay ?? oldText}\n";
        result += $"輸出: {path}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Deletes a hyperlink from a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int?>();
        var cell = arguments?["cell"]?.GetValue<string>();

        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
        {
            throw new ArgumentException("Either hyperlinkIndex or cell is required for delete operation");
        }

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;
        
        int? foundIndex = hyperlinkIndex;
        
        // If cell is provided, find the hyperlink index by cell address
        if (!hyperlinkIndex.HasValue && !string.IsNullOrEmpty(cell))
        {
            foundIndex = null;
            int rowIndex, colIndex;
            try
            {
                CellsHelper.CellNameToIndex(cell, out rowIndex, out colIndex);
            }
            catch
            {
                throw new ArgumentException($"Invalid cell address: {cell}");
            }
            
            for (int i = 0; i < hyperlinks.Count; i++)
            {
                var link = hyperlinks[i];
                var area = link.Area;
                if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                    colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                {
                    foundIndex = i;
                    break;
                }
            }
            
            if (!foundIndex.HasValue)
            {
                throw new ArgumentException($"No hyperlink found at cell {cell}");
            }
        }
        
        if (!foundIndex.HasValue)
        {
            throw new ArgumentException("hyperlinkIndex is required");
        }
        
        int index = foundIndex.Value;
        if (index < 0 || index >= hyperlinks.Count)
        {
            throw new ArgumentException($"超連結索引 {index} 超出範圍 (工作表共有 {hyperlinks.Count} 個超連結)");
        }

        var hyperlink = hyperlinks[index];
        var address = hyperlink.Address ?? "";
        
        hyperlinks.RemoveAt(index);
        workbook.Save(path);
        
        var remainingCount = hyperlinks.Count;
        
        var result = $"成功刪除超連結 #{index}";
        if (!string.IsNullOrEmpty(cell))
        {
            result += $" (單元格: {cell})";
        }
        result += $"\n地址: {address}\n工作表剩餘超連結數: {remainingCount}\n輸出: {path}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Gets all hyperlinks from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all hyperlinks</returns>
    private async Task<string> GetHyperlinksAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的超連結資訊 ===\n");
        result.AppendLine($"總超連結數: {hyperlinks.Count}\n");

        if (hyperlinks.Count == 0)
        {
            result.AppendLine("未找到超連結");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < hyperlinks.Count; i++)
        {
            var hyperlink = hyperlinks[i];
            result.AppendLine($"【超連結 {i}】");
            result.AppendLine($"地址: {hyperlink.Address ?? "(無)"}");
            result.AppendLine($"顯示文字: {hyperlink.TextToDisplay ?? "(無)"}");
            var area = hyperlink.Area;
            result.AppendLine($"位置: 行 {area.StartRow}-{area.EndRow}, 列 {area.StartColumn}-{area.EndColumn}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

