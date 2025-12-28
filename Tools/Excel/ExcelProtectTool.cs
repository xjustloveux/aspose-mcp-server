using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel protection (protect, unprotect, get, set_cell_locked).
///     Merges: ExcelProtectTool, ExcelUnprotectTool, ExcelGetProtectionTool, ExcelProtectWorkbookTool.
/// </summary>
public class ExcelProtectTool : IAsposeTool
{
    private const string OperationProtect = "protect";
    private const string OperationUnprotect = "unprotect";
    private const string OperationGet = "get";
    private const string OperationSetCellLocked = "set_cell_locked";

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        $@"Manage Excel protection. Supports 4 operations: {OperationProtect}, {OperationUnprotect}, {OperationGet}, {OperationSetCellLocked}.

Usage examples:
- Protect sheet: excel_protect(operation='{OperationProtect}', path='book.xlsx', sheetIndex=0, password='password')
- Unprotect sheet: excel_protect(operation='{OperationUnprotect}', path='book.xlsx', sheetIndex=0, password='password')
- Get protection: excel_protect(operation='{OperationGet}', path='book.xlsx', sheetIndex=0)
- Set cell locked: excel_protect(operation='{OperationSetCellLocked}', path='book.xlsx', range='A1:B10', locked=true)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = $@"Operation to perform.
- '{OperationProtect}': Protect workbook or sheet (required params: path, password)
- '{OperationUnprotect}': Unprotect workbook or sheet (required params: path, password)
- '{OperationGet}': Get protection settings (required params: path)
- '{OperationSetCellLocked}': Set cell locked status (required params: path, range, locked)",
                @enum = new[] { OperationProtect, OperationUnprotect, OperationGet, OperationSetCellLocked }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description =
                    "Sheet index (0-based, optional). If not specified, the operation applies to the entire workbook structure instead of a specific worksheet."
            },
            password = new
            {
                type = "string",
                description = "Protection password (required for protect, optional for unprotect)"
            },
            protectWorkbook = new
            {
                type = "boolean",
                description =
                    "Protect workbook structure (optional, for protect operation, default: false). When true with sheetIndex not specified, protects workbook structure. When false with sheetIndex not specified, also protects workbook structure."
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
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLowerInvariant() switch
        {
            OperationProtect => await ProtectAsync(path, outputPath, arguments),
            OperationUnprotect => await UnprotectAsync(path, outputPath, arguments),
            OperationGet => await GetProtectionAsync(path, arguments),
            OperationSetCellLocked => await SetCellLockedAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects workbook or worksheet with password.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex, protectWorkbook, protectionType.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range.</exception>
    private Task<string> ProtectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetString(arguments, "password");
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");
            var protectWorkbook = ArgumentHelper.GetBool(arguments, "protectWorkbook", false);
            var protectStructure = ArgumentHelper.GetBool(arguments, "protectStructure", protectWorkbook);
            var protectWindows = ArgumentHelper.GetBool(arguments, "protectWindows", false);

            using var workbook = new Workbook(path);

            if (protectWorkbook || (!sheetIndex.HasValue && !protectWorkbook))
            {
                var protectionType = ProtectionType.None;
                if (protectStructure && protectWindows)
                    protectionType = ProtectionType.All;
                else if (protectStructure)
                    protectionType = ProtectionType.Structure;
                else if (protectWindows)
                    protectionType = ProtectionType.Windows;

                if (protectionType != ProtectionType.None)
                    workbook.Protect(protectionType, password);
            }
            else if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                worksheet.Protect(ProtectionType.All, password, null);
            }

            workbook.Save(outputPath);

            var target = protectWorkbook ? "workbook" :
                sheetIndex.HasValue ? $"worksheet {sheetIndex.Value}" : "workbook";
            return $"Excel {target} protected with password successfully. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Removes protection from workbook or worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing password, optional sheetIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or password is incorrect.</exception>
    private Task<string> UnprotectAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var password = ArgumentHelper.GetStringNullable(arguments, "password");
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);

            if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);

                if (!worksheet.IsProtected)
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
    ///     Gets protection status for workbook or worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="arguments">JSON arguments containing optional sheetIndex.</param>
    /// <returns>JSON formatted string with protection status.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range.</exception>
    private Task<string> GetProtectionAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetIntNullable(arguments, "sheetIndex");

            using var workbook = new Workbook(path);
            var worksheets = new List<object>();

            if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                worksheets.Add(CreateSheetProtectionInfo(worksheet, sheetIndex.Value));
            }
            else
            {
                for (var i = 0; i < workbook.Worksheets.Count; i++)
                    worksheets.Add(CreateSheetProtectionInfo(workbook.Worksheets[i], i));
            }

            var result = new
            {
                count = worksheets.Count,
                totalWorksheets = workbook.Worksheets.Count,
                worksheets
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Creates protection information object for a worksheet.
    /// </summary>
    /// <param name="worksheet">Worksheet to get protection info from.</param>
    /// <param name="index">Worksheet index (0-based).</param>
    /// <returns>Anonymous object with protection details.</returns>
    private static object CreateSheetProtectionInfo(Worksheet worksheet, int index)
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
    ///     Sets cell locked status.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing range, locked, optional sheetIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range is invalid.</exception>
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