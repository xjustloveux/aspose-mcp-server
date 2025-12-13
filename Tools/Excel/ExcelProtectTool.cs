using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel protection (protect, unprotect, get)
/// Merges: ExcelProtectTool, ExcelUnprotectTool, ExcelGetProtectionTool, ExcelProtectWorkbookTool
/// </summary>
public class ExcelProtectTool : IAsposeTool
{
    public string Description => @"Manage Excel protection. Supports 4 operations: protect, unprotect, get, set_cell_locked.

Usage examples:
- Protect sheet: excel_protect(operation='protect', path='book.xlsx', sheetIndex=0, password='password')
- Unprotect sheet: excel_protect(operation='unprotect', path='book.xlsx', sheetIndex=0, password='password')
- Get protection: excel_protect(operation='get', path='book.xlsx', sheetIndex=0)
- Set cell locked: excel_protect(operation='set_cell_locked', path='book.xlsx', range='A1:B10', locked=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'protect': Protect workbook or sheet (required params: path, password)
- 'unprotect': Unprotect workbook or sheet (required params: path, password)
- 'get': Get protection settings (required params: path)
- 'set_cell_locked': Set cell locked status (required params: path, range, locked)",
                @enum = new[] { "protect", "unprotect", "get", "set_cell_locked" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, protects/unprotects workbook if not specified)"
            },
            password = new
            {
                type = "string",
                description = "Protection password (required for protect, optional for unprotect)"
            },
            protectWorkbook = new
            {
                type = "boolean",
                description = "Protect workbook structure (optional, for protect operation, default: false)"
            },
            protectStructure = new
            {
                type = "boolean",
                description = "Protect workbook structure (optional, for protect operation when protectWorkbook is true, default: true)"
            },
            protectWindows = new
            {
                type = "boolean",
                description = "Protect workbook windows (optional, for protect operation when protectWorkbook is true, default: false)"
            },
            range = new
            {
                type = "string",
                description = "Cell or range (e.g., 'A1' or 'A1:C5', required for set_cell_locked)"
            },
            locked = new
            {
                type = "boolean",
                description = "Locked status (true = locked, false = unlocked, required for set_cell_locked)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");

        return operation.ToLower() switch
        {
            "protect" => await ProtectAsync(arguments, path),
            "unprotect" => await UnprotectAsync(arguments, path),
            "get" => await GetProtectionAsync(arguments, path),
            "set_cell_locked" => await SetCellLockedAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> ProtectAsync(JsonObject? arguments, string path)
    {
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required for protect operation");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();
        var protectWorkbook = arguments?["protectWorkbook"]?.GetValue<bool?>() ?? false;
        var protectStructure = arguments?["protectStructure"]?.GetValue<bool?>() ?? true;
        var protectWindows = arguments?["protectWindows"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);

        if (protectWorkbook || (!sheetIndex.HasValue && !protectWorkbook))
        {
            // Protect workbook with granular control
            var protectionType = ProtectionType.None;
            if (protectStructure && protectWindows)
            {
                protectionType = ProtectionType.All;
            }
            else if (protectStructure)
            {
                protectionType = ProtectionType.Structure;
            }
            else if (protectWindows)
            {
                protectionType = ProtectionType.Windows;
            }
            
            if (protectionType != ProtectionType.None)
            {
                workbook.Protect(protectionType, password);
            }
        }
        else if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
            }
            workbook.Worksheets[sheetIndex.Value].Protect(ProtectionType.All, password, null);
        }

        workbook.Save(path);

        var target = protectWorkbook ? "工作簿" : (sheetIndex.HasValue ? $"工作表 {sheetIndex.Value}" : "工作簿");
        var result = $"Excel {target} protected with password: {path}";
        if (protectWorkbook)
        {
            result += $"\n保護結構: {protectStructure}\n保護視窗: {protectWindows}";
        }
        return await Task.FromResult(result);
    }

    private async Task<string> UnprotectAsync(JsonObject? arguments, string path)
    {
        var password = arguments?["password"]?.GetValue<string>();
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
            }

            var worksheet = workbook.Worksheets[sheetIndex.Value];
            var wasProtected = worksheet.IsProtected;
            worksheet.Unprotect(password);

            workbook.Save(path);
            return await Task.FromResult($"工作表解除保護完成: {worksheet.Name}\n原狀態: {(wasProtected ? "已保護" : "未保護")}\n輸出: {path}");
        }
        else
        {
            workbook.Unprotect(password);
            workbook.Save(path);
            return await Task.FromResult($"工作簿保護已解除\n輸出: {path}");
        }
    }

    private async Task<string> GetProtectionAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine("=== Excel 保護設定資訊 ===\n");

        result.AppendLine("【工作簿保護】");
        result.AppendLine("注意: 工作簿保護狀態需要通過保護方法檢查");
        result.AppendLine();

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"工作表索引 {sheetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }
            AppendSheetProtection(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value);
        }
        else
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                AppendSheetProtection(result, workbook.Worksheets[i], i);
                if (i < workbook.Worksheets.Count - 1) result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSheetProtection(StringBuilder result, Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        result.AppendLine($"【工作表 {index}: {worksheet.Name}】");
        result.AppendLine($"保護狀態: {(protection.IsProtectedWithPassword ? "已保護" : "未保護")}");
        result.AppendLine($"允許選擇鎖定單元格: {protection.AllowSelectingLockedCell}");
        result.AppendLine($"允許選擇未鎖定單元格: {protection.AllowSelectingUnlockedCell}");
        result.AppendLine($"允許格式化單元格: {protection.AllowFormattingCell}");
        result.AppendLine($"允許格式化列: {protection.AllowFormattingColumn}");
        result.AppendLine($"允許格式化行: {protection.AllowFormattingRow}");
        result.AppendLine($"允許插入列: {protection.AllowInsertingColumn}");
        result.AppendLine($"允許插入行: {protection.AllowInsertingRow}");
        result.AppendLine($"允許插入超連結: {protection.AllowInsertingHyperlink}");
        result.AppendLine($"允許刪除列: {protection.AllowDeletingColumn}");
        result.AppendLine($"允許刪除行: {protection.AllowDeletingRow}");
        result.AppendLine($"允許排序: {protection.AllowSorting}");
        result.AppendLine($"允許自動篩選: {protection.AllowFiltering}");
        result.AppendLine($"允許使用樞紐表: {protection.AllowUsingPivotTable}");
        result.AppendLine($"允許編輯對象: {protection.AllowEditingObject}");
        result.AppendLine($"允許編輯場景: {protection.AllowEditingScenario}");
    }

    private async Task<string> SetCellLockedAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for set_cell_locked operation");
        var locked = arguments?["locked"]?.GetValue<bool>() ?? throw new ArgumentException("locked is required for set_cell_locked operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var style = workbook.CreateStyle();
        style.IsLocked = locked;

        var styleFlag = new StyleFlag { Locked = true };
        cellRange.ApplyStyle(style, styleFlag);

        workbook.Save(path);
        return await Task.FromResult($"Cell lock status set to {(locked ? "locked" : "unlocked")} for range {range} in sheet {sheetIndex}: {path}");
    }
}
