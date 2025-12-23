using System.Text;
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

        return operation.ToLower() switch
        {
            "protect" => await ProtectAsync(arguments, path),
            "unprotect" => await UnprotectAsync(arguments, path),
            "get" => await GetProtectionAsync(arguments, path),
            "set_cell_locked" => await SetCellLockedAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects workbook or worksheet with password
    /// </summary>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex, protectWorkbook, protectionType</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message</returns>
    private Task<string> ProtectAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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
            var result = $"Excel {target} protected with password: {outputPath}";
            if (protectWorkbook)
                result += $"\nProtect structure: {protectStructure}\nProtect windows: {protectWindows}";
            return result;
        });
    }

    /// <summary>
    ///     Removes protection from workbook or worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message</returns>
    private Task<string> UnprotectAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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
                    return $"Worksheet '{worksheet.Name}' is not protected. No password needed.\nOutput: {outputPath}";
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
                return $"Worksheet protection removed: {worksheet.Name}\nOutput: {outputPath}";
            }

            workbook.Unprotect(password);
            workbook.Save(outputPath);
            return $"Workbook protection removed\nOutput: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets protection status for workbook or worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional sheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with protection status</returns>
    private Task<string> GetProtectionAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);
            var result = new StringBuilder();

            result.AppendLine("=== Excel Protection Settings Information ===\n");

            result.AppendLine("[Workbook Protection]");
            result.AppendLine("Note: Workbook protection status needs to be checked through protection methods");
            result.AppendLine();

            if (sheetIndex.HasValue)
            {
                if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
                    throw new ArgumentException(
                        $"Worksheet index {sheetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");
                AppendSheetProtection(result, workbook.Worksheets[sheetIndex.Value], sheetIndex.Value);
            }
            else
            {
                for (var i = 0; i < workbook.Worksheets.Count; i++)
                {
                    AppendSheetProtection(result, workbook.Worksheets[i], i);
                    if (i < workbook.Worksheets.Count - 1) result.AppendLine();
                }
            }

            return result.ToString();
        });
    }

    private void AppendSheetProtection(StringBuilder result, Worksheet worksheet, int index)
    {
        var protection = worksheet.Protection;
        result.AppendLine($"[Worksheet {index}: {worksheet.Name}]");
        result.AppendLine($"Protection status: {(protection.IsProtectedWithPassword ? "Protected" : "Not protected")}");
        result.AppendLine($"Allow selecting locked cells: {protection.AllowSelectingLockedCell}");
        result.AppendLine($"Allow selecting unlocked cells: {protection.AllowSelectingUnlockedCell}");
        result.AppendLine($"Allow formatting cells: {protection.AllowFormattingCell}");
        result.AppendLine($"Allow formatting columns: {protection.AllowFormattingColumn}");
        result.AppendLine($"Allow formatting rows: {protection.AllowFormattingRow}");
        result.AppendLine($"Allow inserting columns: {protection.AllowInsertingColumn}");
        result.AppendLine($"Allow inserting rows: {protection.AllowInsertingRow}");
        result.AppendLine($"Allow inserting hyperlinks: {protection.AllowInsertingHyperlink}");
        result.AppendLine($"Allow deleting columns: {protection.AllowDeletingColumn}");
        result.AppendLine($"Allow deleting rows: {protection.AllowDeletingRow}");
        result.AppendLine($"Allow sorting: {protection.AllowSorting}");
        result.AppendLine($"Allow auto filtering: {protection.AllowFiltering}");
        result.AppendLine($"Allow using pivot tables: {protection.AllowUsingPivotTable}");
        result.AppendLine($"Allow editing objects: {protection.AllowEditingObject}");
        result.AppendLine($"Allow editing scenarios: {protection.AllowEditingScenario}");
    }

    /// <summary>
    ///     Sets cell locked status
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, isLocked, optional sheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message</returns>
    private Task<string> SetCellLockedAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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
                $"Cell lock status set to {(locked ? "locked" : "unlocked")} for range {range} in sheet {sheetIndex}: {outputPath}";
        });
    }
}