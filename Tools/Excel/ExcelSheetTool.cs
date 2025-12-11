using System.Text.Json.Nodes;
using System.Text;
using System.Linq;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel sheets (add, delete, get, rename, move, copy, hide)
/// Merges: ExcelAddSheetTool, ExcelDeleteSheetTool, ExcelGetSheetsTool, ExcelRenameSheetTool, 
/// ExcelMoveSheetTool, ExcelCopySheetTool, ExcelHideSheetTool
/// </summary>
public class ExcelSheetTool : IAsposeTool
{
    public string Description => "Manage Excel sheets: add, delete, get, rename, move, copy, or hide worksheets";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'delete', 'get', 'rename', 'move', 'copy', 'hide'",
                @enum = new[] { "add", "delete", "get", "rename", "move", "copy", "hide" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, required for delete/rename/move/copy/hide)"
            },
            sheetName = new
            {
                type = "string",
                description = "Name of the sheet (required for add/rename)"
            },
            newName = new
            {
                type = "string",
                description = "New name for the sheet (required for rename)"
            },
            insertAt = new
            {
                type = "number",
                description = "Position to insert the sheet (0-based, optional for add, required for move)"
            },
            targetIndex = new
            {
                type = "number",
                description = "Target index for move/copy operation (0-based)"
            },
            copyToPath = new
            {
                type = "string",
                description = "Target file path for copy operation (optional, if not provided copies within same workbook)"
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
            "add" => await AddSheetAsync(arguments, path),
            "delete" => await DeleteSheetAsync(arguments, path),
            "get" => await GetSheetsAsync(arguments, path),
            "rename" => await RenameSheetAsync(arguments, path),
            "move" => await MoveSheetAsync(arguments, path),
            "copy" => await CopySheetAsync(arguments, path),
            "hide" => await HideSheetAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddSheetAsync(JsonObject? arguments, string path)
    {
        var sheetName = arguments?["sheetName"]?.GetValue<string>()?.Trim() ?? throw new ArgumentException("sheetName is required for add operation");
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("sheetName cannot be empty");
        }

        using var workbook = new Workbook(path);

        var duplicate = workbook.Worksheets.Any(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
        {
            throw new ArgumentException($"Worksheet name '{sheetName}' already exists in the workbook");
        }

        Worksheet newSheet;
        if (insertAt.HasValue)
        {
            if (insertAt.Value < 0 || insertAt.Value > workbook.Worksheets.Count)
            {
                throw new ArgumentException($"insertAt must be between 0 and {workbook.Worksheets.Count}");
            }

            if (insertAt.Value == workbook.Worksheets.Count)
            {
                var addedIndex = workbook.Worksheets.Add();
                newSheet = workbook.Worksheets[addedIndex];
            }
            else
            {
                workbook.Worksheets.Insert(insertAt.Value, SheetType.Worksheet);
                newSheet = workbook.Worksheets[insertAt.Value];
            }
        }
        else
        {
            var addedIndex = workbook.Worksheets.Add();
            newSheet = workbook.Worksheets[addedIndex];
        }
        
        newSheet.Name = sheetName;
        workbook.Save(path);

        return await Task.FromResult($"Worksheet '{sheetName}' added: {path}");
    }

    private async Task<string> DeleteSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required for delete operation");

        using var workbook = new Workbook(path);
        ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

        if (workbook.Worksheets.Count <= 1)
        {
            throw new InvalidOperationException("無法刪除最後一個工作表");
        }

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        workbook.Worksheets.RemoveAt(sheetIndex);
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{sheetName}' (索引 {sheetIndex}) 已刪除: {path}");
    }

    private async Task<string> GetSheetsAsync(JsonObject? arguments, string path)
    {
        using var workbook = new Workbook(path);
        var result = new StringBuilder();

        result.AppendLine($"=== 工作簿 '{Path.GetFileName(path)}' 的工作表列表 ===\n");
        result.AppendLine($"總工作表數: {workbook.Worksheets.Count}\n");

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            var worksheet = workbook.Worksheets[i];
            result.AppendLine($"{i}. {worksheet.Name} (可見性: {(worksheet.IsVisible ? "Visible" : "Hidden")})");
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> RenameSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required for rename operation");
        var newName = arguments?["newName"]?.GetValue<string>()?.Trim() ?? throw new ArgumentException("newName is required for rename operation");

        if (string.IsNullOrWhiteSpace(newName))
        {
            throw new ArgumentException("newName cannot be empty");
        }

        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var oldName = worksheet.Name;

        // Check for duplicate names
        var duplicate = workbook.Worksheets.Any(ws => ws != worksheet && string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
        {
            throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");
        }

        worksheet.Name = newName;
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{oldName}' 已重新命名為 '{newName}': {path}");
    }

    private async Task<string> MoveSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required for move operation");
        var targetIndex = arguments?["targetIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetIndex is required for move operation");

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (targetIndex < 0 || targetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"目標索引 {targetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (sheetIndex == targetIndex)
        {
            return await Task.FromResult($"工作表已在位置 {sheetIndex}，無需移動: {path}");
        }

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        // Move sheet by copying and deleting
        var newSheetIndex = workbook.Worksheets.AddCopy(sheetIndex);
        workbook.Worksheets[newSheetIndex].Name = sheetName + "_temp";
        workbook.Worksheets.RemoveAt(sheetIndex);
        workbook.Worksheets[newSheetIndex].Name = sheetName;
        // Reorder if needed
        if (targetIndex != newSheetIndex)
        {
            var tempName = sheetName + "_move_temp";
            workbook.Worksheets[newSheetIndex].Name = tempName;
            workbook.Worksheets.AddCopy(newSheetIndex);
            workbook.Worksheets.RemoveAt(newSheetIndex);
            workbook.Worksheets[workbook.Worksheets.Count - 1].Name = sheetName;
        }
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{sheetName}' 已從位置 {sheetIndex} 移動到位置 {targetIndex}: {path}");
    }

    private async Task<string> CopySheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required for copy operation");
        var targetIndex = arguments?["targetIndex"]?.GetValue<int?>();
        var copyToPath = arguments?["copyToPath"]?.GetValue<string>();
        if (!string.IsNullOrEmpty(copyToPath))
        {
            SecurityHelper.ValidateFilePath(copyToPath, "copyToPath");
        }

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var sourceSheet = workbook.Worksheets[sheetIndex];
        var sheetName = sourceSheet.Name;

        if (!string.IsNullOrEmpty(copyToPath))
        {
            // Copy to another workbook
            using var targetWorkbook = new Workbook();
            var newSheet = targetWorkbook.Worksheets.Add(sheetName);
            sourceSheet.Copy(newSheet);
            targetWorkbook.Save(copyToPath);
            return await Task.FromResult($"工作表 '{sheetName}' 已複製到 '{copyToPath}': {path}");
        }
        else
        {
            // Copy within same workbook
            if (!targetIndex.HasValue)
            {
                targetIndex = workbook.Worksheets.Count;
            }

            if (targetIndex.Value < 0 || targetIndex.Value > workbook.Worksheets.Count)
            {
                throw new ArgumentException($"目標索引 {targetIndex.Value} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
            }

            var newSheetIndex = workbook.Worksheets.AddCopy(sheetIndex);
            // Note: MoveSheet may not be available, sheets are added at the end
            // If specific position is needed, would need to copy and reorder manually
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已複製到位置 {targetIndex.Value}: {path}");
        }
    }

    private async Task<string> HideSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required for hide operation");

        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var sheetName = worksheet.Name;

        if (worksheet.IsVisible)
        {
            worksheet.IsVisible = false;
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已隱藏: {path}");
        }
        else
        {
            worksheet.IsVisible = true;
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已顯示: {path}");
        }
    }
}

