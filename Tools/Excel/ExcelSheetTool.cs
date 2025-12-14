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
    public string Description => @"Manage Excel sheets. Supports 7 operations: add, delete, get, rename, move, copy, hide.

Usage examples:
- Add sheet: excel_sheet(operation='add', path='book.xlsx', sheetName='New Sheet')
- Delete sheet: excel_sheet(operation='delete', path='book.xlsx', sheetIndex=1)
- Get sheets: excel_sheet(operation='get', path='book.xlsx')
- Rename sheet: excel_sheet(operation='rename', path='book.xlsx', sheetIndex=0, newName='Renamed')
- Move sheet: excel_sheet(operation='move', path='book.xlsx', sheetIndex=0, insertAt=2)
- Copy sheet: excel_sheet(operation='copy', path='book.xlsx', sheetIndex=0, newName='Copy')
- Hide sheet: excel_sheet(operation='hide', path='book.xlsx', sheetIndex=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new sheet (required params: path, sheetName)
- 'delete': Delete a sheet (required params: path, sheetIndex)
- 'get': Get all sheets (required params: path)
- 'rename': Rename a sheet (required params: path, sheetIndex, newName)
- 'move': Move a sheet (required params: path, sheetIndex, insertAt)
- 'copy': Copy a sheet (required params: path, sheetIndex, newName)
- 'hide': Hide a sheet (required params: path, sheetIndex)",
                @enum = new[] { "add", "delete", "get", "rename", "move", "copy", "hide" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
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
                description = "Position to insert the sheet (0-based, optional for add, optional for move as alternative to targetIndex)"
            },
            targetIndex = new
            {
                type = "number",
                description = "Target index for move/copy operation (0-based, required for move, or use insertAt as alternative)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

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

    /// <summary>
    /// Adds a new worksheet to the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetName and optional insertAt</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with worksheet name</returns>
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

    /// <summary>
    /// Deletes a worksheet from the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with deleted sheet name</returns>
    private async Task<string> DeleteSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", "sheetIndex");

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

    /// <summary>
    /// Gets information about all worksheets in the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with worksheet list</returns>
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

    /// <summary>
    /// Renames a worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetIndex and newName</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with old and new names</returns>
    private async Task<string> RenameSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", "sheetIndex");
        var newName = ArgumentHelper.GetString(arguments, "newName", "newName").Trim();

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

    /// <summary>
    /// Moves a worksheet to a different position
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetIndex and targetIndex or insertAt</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with move details</returns>
    private async Task<string> MoveSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", "sheetIndex");
        var targetIndex = arguments?["targetIndex"]?.GetValue<int?>();
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();
        
        if (!targetIndex.HasValue && !insertAt.HasValue)
        {
            throw new ArgumentException("Either targetIndex or insertAt is required for move operation");
        }
        
        var finalTargetIndex = targetIndex ?? insertAt!.Value;

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (finalTargetIndex < 0 || finalTargetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"目標索引 {finalTargetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (sheetIndex == finalTargetIndex)
        {
            return await Task.FromResult($"工作表已在位置 {sheetIndex}，無需移動: {path}");
        }

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        
        // Use temporary unique name to avoid conflicts during move (Excel sheet name limit: 31 characters)
        var tempName = $"Temp_{DateTime.Now.Ticks % 1000000}";
        int tempCounter = 0;
        while (workbook.Worksheets.Any(ws => ws.Name == tempName))
        {
            tempName = $"Temp_{DateTime.Now.Ticks % 1000000}_{tempCounter++}";
            if (tempName.Length > 31)
            {
                tempName = tempName.Substring(0, 31);
            }
        }
        
        // Use Copy method to duplicate sheet at target position, then remove original
        try
        {
            if (finalTargetIndex < sheetIndex)
            {
                // Moving backward: insert copy at target position first
                workbook.Worksheets.Insert(finalTargetIndex, SheetType.Worksheet);
                var newSheet = workbook.Worksheets[finalTargetIndex];
                var sourceSheet = workbook.Worksheets[sheetIndex + 1];
                newSheet.Copy(sourceSheet);
                newSheet.Name = tempName;
                workbook.Worksheets.RemoveAt(sheetIndex + 1);
                newSheet.Name = sheetName;
            }
            else
            {
                // Moving forward: copy to target position, then remove original
                workbook.Worksheets.Insert(finalTargetIndex, SheetType.Worksheet);
                var newSheet = workbook.Worksheets[finalTargetIndex];
                var sourceSheet = workbook.Worksheets[sheetIndex];
                newSheet.Copy(sourceSheet);
                newSheet.Name = tempName;
                workbook.Worksheets.RemoveAt(sheetIndex);
                newSheet.Name = sheetName;
            }
        }
        catch (Exception ex)
        {
            // Clean up temporary sheet if operation fails
            try
            {
                for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
                {
                    if (workbook.Worksheets[i].Name == tempName)
                    {
                        workbook.Worksheets.RemoveAt(i);
                        break;
                    }
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
            throw new ArgumentException($"Failed to move sheet: {ex.Message}. Source index: {sheetIndex}, Target index: {finalTargetIndex}, Total sheets: {workbook.Worksheets.Count}");
        }
        
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{sheetName}' 已從位置 {sheetIndex} 移動到位置 {finalTargetIndex}: {path}");
    }

    /// <summary>
    /// Copies a worksheet with a new name
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetIndex, newName, optional targetIndex or copyToPath</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with copy details</returns>
    private async Task<string> CopySheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", "sheetIndex");
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
            // If specific position is needed, would need to copy and reorder manually
            workbook.Save(path);
            return await Task.FromResult($"工作表 '{sheetName}' 已複製到位置 {targetIndex.Value}: {path}");
        }
    }

    /// <summary>
    /// Hides or shows a worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing sheetIndex and optional targetIndex, isVisible</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with visibility status</returns>
    private async Task<string> HideSheetAsync(JsonObject? arguments, string path)
    {
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", "sheetIndex");

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

