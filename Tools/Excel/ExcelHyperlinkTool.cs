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
    public string Description => "Manage Excel hyperlinks: add, edit, delete, or get hyperlinks from cells";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'delete', 'get'",
                @enum = new[] { "add", "edit", "delete", "get" }
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
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1', required for add)"
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
                description = "Hyperlink index (0-based, required for edit/delete)"
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
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
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

    private async Task<string> AddHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for add operation");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required for add operation");
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

    private async Task<string> EditHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required for edit operation");
        var address = arguments?["address"]?.GetValue<string>();
        var textToDisplay = arguments?["textToDisplay"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;
        
        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinks.Count)
        {
            throw new ArgumentException($"超連結索引 {hyperlinkIndex} 超出範圍 (工作表共有 {hyperlinks.Count} 個超連結)");
        }

        var hyperlink = hyperlinks[hyperlinkIndex];
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
        
        var result = $"成功編輯超連結 #{hyperlinkIndex}\n";
        result += $"舊地址: {oldAddress}\n";
        result += $"新地址: {hyperlink.Address ?? oldAddress}\n";
        result += $"舊顯示文字: {oldText}\n";
        result += $"新顯示文字: {hyperlink.TextToDisplay ?? oldText}\n";
        result += $"輸出: {path}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required for delete operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;
        
        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinks.Count)
        {
            throw new ArgumentException($"超連結索引 {hyperlinkIndex} 超出範圍 (工作表共有 {hyperlinks.Count} 個超連結)");
        }

        var hyperlink = hyperlinks[hyperlinkIndex];
        var address = hyperlink.Address ?? "";
        
        hyperlinks.RemoveAt(hyperlinkIndex);
        workbook.Save(path);
        
        var remainingCount = hyperlinks.Count;
        
        return await Task.FromResult($"成功刪除超連結 #{hyperlinkIndex}\n地址: {address}\n工作表剩餘超連結數: {remainingCount}\n輸出: {path}");
    }

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

