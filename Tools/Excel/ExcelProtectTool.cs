using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel protection (protect, unprotect, get)
///     Merges: ExcelProtectTool, ExcelUnprotectTool, ExcelGetProtectionTool, ExcelProtectWorkbookTool
/// </summary>
public class ExcelProtectTool : IAsposeTool
{
    public string Description =>
        @"Manage Excel protection. Supports 4 operations: protect, unprotect, get, set_cell_locked.

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
                description =
                    "Protect workbook structure (optional, for protect operation when protectWorkbook is true, default: true)"
            },
            protectWindows = new
            {
                type = "boolean",
                description =
                    "Protect workbook windows (optional, for protect operation when protectWorkbook is true, default: false)"
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
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for protect/unprotect/set_cell_locked operations, defaults to input path)"
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

        return operation.ToLower() switch
        {
            "protect" => await ProtectAsync(path, outputPath, arguments),
            "unprotect" => await UnprotectAsync(path, outputPath, arguments),
            "get" => await GetProtectionAsync(path, arguments),
            "set_cell_locked" => await SetCellLockedAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects workbook or worksheet with password
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex, protectWorkbook, protectionType</param>
    /// <returns>Success message</returns>
    private Task<string> ProtectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetString(arguments, "password");
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
            var protectWorkbook = ArgumentHelper.GetBool(arguments, "protectWorkbook", false);
            // protectStructure defaults to true when protectWorkbook is true, otherwise defaults to false
            var protectStructure = ArgumentHelper.GetBool(arguments, "protectStructure", protectWorkbook);
            var protectWindows = ArgumentHelper.GetBool(arguments, "protectWindows", false);

            using var workbook = new Workbook(path);

            if (protectWorkbook || (!sheetIndex.HasValue && !protectWorkbook))
            {
                // Protect workbook with granular control
                var protectionType = ProtectionType.None;
                if (protectStructure && protectWindows)
                    protectionType = ProtectionType.All;
                else if (protectStructure)
                    protectionType = ProtectionType.Structure;
                else if (protectWindows) protectionType = ProtectionType.Windows;

                if (protectionType != ProtectionType.None) workbook.Protect(protectionType, password);
            }
            else if (sheetIndex.HasValue)
            {
                if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                    throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
                workbook.Worksheets[sheetIndex.Value].Protect(ProtectionType.All, password, null);
            }

            workbook.Save(outputPath);

            var target = protectWorkbook ? "workbook" :
                sheetIndex.HasValue ? $"worksheet {sheetIndex.Value}" : "workbook";
            return $"Excel {target} protected with password successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes protection from workbook or worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex</param>
    /// <returns>Success message</returns>
    private Task<string> UnprotectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetStringNullable(arguments, "password");
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);

            if (sheetIndex.HasValue)
            {
                if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                    throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");

                var worksheet = workbook.Worksheets[sheetIndex.Value];
                var wasProtected = worksheet.IsProtected;

                if (!wasProtected)
                {
                    workbook.Save(outputPath);
                    return $"Worksheet '{worksheet.Name}' is not protected. Output: {outputPath}";
                }

                try
                {
                    worksheet.Unprotect(password);
                }
                catch (Exception ex)
                {
                    workbook.Save(outputPath);
                    throw new ArgumentException(
                        $"Incorrect password. Cannot unprotect worksheet '{worksheet.Name}'. Error: {ex.Message}");
                }

                workbook.Save(outputPath);
                return $"Worksheet '{worksheet.Name}' protection removed successfully. Output: {outputPath}";
            }

            workbook.Unprotect(password);
            workbook.Save(outputPath);
            return $"Workbook protection removed successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets protection status for workbook or worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="arguments">JSON arguments containing optional sheetIndex</param>
    /// <returns>JSON formatted string with protection status</returns>
    private Task<string> GetProtectionAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);
            var worksheets = new List<object>();

            if (sheetIndex.HasValue)
            {
                if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                    throw new ArgumentException(
                        $"Worksheet index {sheetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");
                worksheets.Add(GetSheetProtectionInfo(workbook.Worksheets[sheetIndex.Value], sheetIndex.Value));
            }
            else
            {
                for (var i = 0; i < workbook.Worksheets.Count; i++)
                    worksheets.Add(GetSheetProtectionInfo(workbook.Worksheets[i], i));
            }

            var result = new
            {
                count = worksheets.Count,
                worksheets
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets protection information for a worksheet
    /// </summary>
    /// <param name="worksheet">Worksheet to get protection info from</param>
    /// <param name="index">Worksheet index</param>
    /// <returns>Anonymous object with protection details</returns>
    private object GetSheetProtectionInfo(Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        return new
        {
            index,
            name = worksheet.Name,
            isProtected = protection.IsProtectedWithPassword,
            allowSelectingLockedCell = protection.AllowSelectingLockedCell,
            allowSelectingUnlockedCell = protection.AllowSelectingUnlockedCell,
            allowFormattingCell = protection.AllowFormattingCell,
            allowFormattingColumn = protection.AllowFormattingColumn,
            allowFormattingRow = protection.AllowFormattingRow,
            allowInsertingColumn = protection.AllowInsertingColumn,
            allowInsertingRow = protection.AllowInsertingRow,
            allowInsertingHyperlink = protection.AllowInsertingHyperlink,
            allowDeletingColumn = protection.AllowDeletingColumn,
            allowDeletingRow = protection.AllowDeletingRow,
            allowSorting = protection.AllowSorting,
            allowFiltering = protection.AllowFiltering,
            allowUsingPivotTable = protection.AllowUsingPivotTable,
            allowEditingObject = protection.AllowEditingObject,
            allowEditingScenario = protection.AllowEditingScenario
        };
    }

    /// <summary>
    ///     Sets cell locked status
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing range, isLocked, optional sheetIndex</param>
    /// <returns>Success message</returns>
    private Task<string> SetCellLockedAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);
            var range = ArgumentHelper.GetString(arguments, "range");
            var locked = ArgumentHelper.GetBool(arguments, "locked", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            var style = workbook.CreateStyle();
            style.IsLocked = locked;

            var styleFlag = new StyleFlag { Locked = true };
            cellRange.ApplyStyle(style, styleFlag);

            workbook.Save(outputPath);
            return
                $"Cell lock status set to {(locked ? "locked" : "unlocked")} for range {range} in sheet {sheetIndex}. Output: {outputPath}";
        });
    }
}