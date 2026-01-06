using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel protection (protect, unprotect, get, set_cell_locked).
///     Merges: ExcelProtectTool, ExcelUnprotectTool, ExcelGetProtectionTool, ExcelProtectWorkbookTool.
/// </summary>
[McpServerToolType]
public class ExcelProtectTool
{
    /// <summary>
    ///     Operation name for protecting workbook or sheet.
    /// </summary>
    private const string OperationProtect = "protect";

    /// <summary>
    ///     Operation name for removing protection.
    /// </summary>
    private const string OperationUnprotect = "unprotect";

    /// <summary>
    ///     Operation name for getting protection status.
    /// </summary>
    private const string OperationGet = "get";

    /// <summary>
    ///     Operation name for setting cell locked status.
    /// </summary>
    private const string OperationSetCellLocked = "set_cell_locked";

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelProtectTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelProtectTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel protection operation (protect, unprotect, get, set_cell_locked).
    /// </summary>
    /// <param name="operation">The operation to perform: protect, unprotect, get, set_cell_locked.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">
    ///     Sheet index (0-based, optional). If not specified, the operation applies to the entire
    ///     workbook structure.
    /// </param>
    /// <param name="password">Protection password (required for protect, optional for unprotect).</param>
    /// <param name="protectWorkbook">Protect workbook structure (optional, for protect operation, default: false).</param>
    /// <param name="protectStructure">
    ///     Protect workbook structure (optional, for protect operation when protectWorkbook is
    ///     true, default: true).
    /// </param>
    /// <param name="protectWindows">
    ///     Protect workbook windows (optional, for protect operation when protectWorkbook is true,
    ///     default: false).
    /// </param>
    /// <param name="range">Cell or range (e.g., 'A1' or 'A1:C5', required for set_cell_locked).</param>
    /// <param name="locked">Locked status (true = locked, false = unlocked, required for set_cell_locked).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_protect")]
    [Description(@"Manage Excel protection. Supports 4 operations: protect, unprotect, get, set_cell_locked.

Usage examples:
- Protect sheet: excel_protect(operation='protect', path='book.xlsx', sheetIndex=0, password='password')
- Unprotect sheet: excel_protect(operation='unprotect', path='book.xlsx', sheetIndex=0, password='password')
- Get protection: excel_protect(operation='get', path='book.xlsx', sheetIndex=0)
- Set cell locked: excel_protect(operation='set_cell_locked', path='book.xlsx', range='A1:B10', locked=true)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'protect': Protect workbook or sheet (required params: path, password)
- 'unprotect': Unprotect workbook or sheet (required params: path, password)
- 'get': Get protection settings (required params: path)
- 'set_cell_locked': Set cell locked status (required params: path, range, locked)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description(
            "Sheet index (0-based, optional). If not specified, the operation applies to the entire workbook structure.")]
        int? sheetIndex = null,
        [Description("Protection password (required for protect, optional for unprotect)")]
        string? password = null,
        [Description("Protect workbook structure (optional, for protect operation, default: false)")]
        bool protectWorkbook = false,
        [Description(
            "Protect workbook structure (optional, for protect operation when protectWorkbook is true, default: true)")]
        bool protectStructure = true,
        [Description(
            "Protect workbook windows (optional, for protect operation when protectWorkbook is true, default: false)")]
        bool protectWindows = false,
        [Description("Cell or range (e.g., 'A1' or 'A1:C5', required for set_cell_locked)")]
        string? range = null,
        [Description("Locked status (true = locked, false = unlocked, required for set_cell_locked)")]
        bool locked = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLowerInvariant() switch
        {
            OperationProtect => Protect(ctx, outputPath, sheetIndex, password, protectWorkbook, protectStructure,
                protectWindows),
            OperationUnprotect => Unprotect(ctx, outputPath, sheetIndex, password),
            OperationGet => GetProtection(ctx, sheetIndex),
            OperationSetCellLocked => SetCellLocked(ctx, outputPath, sheetIndex ?? 0, range, locked),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Protects workbook or worksheet with password.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based), or null for workbook-level protection.</param>
    /// <param name="password">The protection password.</param>
    /// <param name="protectWorkbook">Whether to protect the workbook structure.</param>
    /// <param name="protectStructure">Whether to protect the workbook structure (when protectWorkbook is true).</param>
    /// <param name="protectWindows">Whether to protect the workbook windows (when protectWorkbook is true).</param>
    /// <returns>A message indicating the result of the protection operation.</returns>
    /// <exception cref="ArgumentException">Thrown when password is null or empty.</exception>
    private static string Protect(DocumentContext<Workbook> ctx, string? outputPath, int? sheetIndex,
        string? password, bool protectWorkbook, bool protectStructure, bool protectWindows)
    {
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("password is required for protect operation");

        var workbook = ctx.Document;

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

        ctx.Save(outputPath);

        var target = protectWorkbook ? "workbook" :
            sheetIndex.HasValue ? $"worksheet {sheetIndex.Value}" : "workbook";
        return $"Excel {target} protected with password successfully. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Removes protection from workbook or worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based), or null for workbook-level unprotection.</param>
    /// <param name="password">The protection password used to unprotect.</param>
    /// <returns>A message indicating the result of the unprotection operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the password is incorrect.</exception>
    private static string Unprotect(DocumentContext<Workbook> ctx, string? outputPath, int? sheetIndex,
        string? password)
    {
        var workbook = ctx.Document;

        if (sheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);

            if (!worksheet.IsProtected)
            {
                ctx.Save(outputPath);
                return $"Worksheet '{worksheet.Name}' is not protected. {ctx.GetOutputMessage(outputPath)}";
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

            ctx.Save(outputPath);
            return $"Worksheet '{worksheet.Name}' protection removed successfully. {ctx.GetOutputMessage(outputPath)}";
        }

        workbook.Unprotect(password);
        ctx.Save(outputPath);
        return $"Workbook protection removed successfully. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets protection status for workbook or worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The sheet index (0-based), or null to get all worksheets' protection status.</param>
    /// <returns>A JSON string containing the protection status information.</returns>
    private static string GetProtection(DocumentContext<Workbook> ctx, int? sheetIndex)
    {
        var workbook = ctx.Document;
        List<object> worksheets = [];

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
    }

    /// <summary>
    ///     Creates protection information object for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get protection information from.</param>
    /// <param name="index">The index of the worksheet.</param>
    /// <returns>An anonymous object containing the worksheet's protection details.</returns>
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
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="range">The cell range to set locked status (e.g., 'A1:B10').</param>
    /// <param name="locked">Whether the cells should be locked.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range is null or empty.</exception>
    private static string SetCellLocked(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, bool locked)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for set_cell_locked operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        var style = workbook.CreateStyle();
        style.IsLocked = locked;

        var styleFlag = new StyleFlag { Locked = true };
        cellRange.ApplyStyle(style, styleFlag);

        ctx.Save(outputPath);
        return
            $"Cell lock status set to {(locked ? "locked" : "unlocked")} for range {range} in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}