using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel freeze panes (freeze/unfreeze/get).
/// </summary>
[McpServerToolType]
public class ExcelFreezePanesTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelFreezePanesTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelFreezePanesTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_freeze_panes")]
    [Description(@"Manage Excel freeze panes. Supports 3 operations: freeze, unfreeze, get.

Usage examples:
- Freeze panes: excel_freeze_panes(operation='freeze', path='book.xlsx', row=1, column=1)
- Unfreeze panes: excel_freeze_panes(operation='unfreeze', path='book.xlsx')
- Get freeze status: excel_freeze_panes(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'freeze': Freeze panes at specified row and column (required params: path, row, column)
- 'unfreeze': Remove freeze panes (required params: path)
- 'get': Get current freeze panes status (required params: path)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Number of rows to freeze from top (0-based, required for freeze)")]
        int row = 0,
        [Description("Number of columns to freeze from left (0-based, required for freeze)")]
        int column = 0)
    {
        return operation.ToLower() switch
        {
            "freeze" => FreezePanes(path, sessionId, outputPath, sheetIndex, row, column),
            "unfreeze" => UnfreezePanes(path, sessionId, outputPath, sheetIndex),
            "get" => GetFreezePanes(path, sessionId, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Freezes panes at the specified row and column.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="row">The number of rows to freeze from top.</param>
    /// <param name="column">The number of columns to freeze from left.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string FreezePanes(string? path, string? sessionId, string? outputPath, int sheetIndex, int row, int column)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        worksheet.FreezePanes(row + 1, column + 1, row, column);

        ctx.Save(outputPath);
        return $"Frozen panes at row {row}, column {column}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Removes freeze panes from the worksheet.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string UnfreezePanes(string? path, string? sessionId, string? outputPath, int sheetIndex)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        worksheet.UnFreezePanes();

        ctx.Save(outputPath);
        return $"Unfrozen panes. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets the current freeze panes status.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the freeze panes status information.</returns>
    private string GetFreezePanes(string? path, string? sessionId, int sheetIndex)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);

        var isFrozen = worksheet.PaneState == PaneStateType.Frozen;
        int? frozenRow = null;
        int? frozenColumn = null;
        int? frozenRows = null;
        int? frozenColumns = null;

        if (isFrozen)
        {
            worksheet.GetFreezedPanes(out var r, out var col, out var rows, out var cols);
            frozenRow = r > 0 ? r - 1 : 0;
            frozenColumn = col > 0 ? col - 1 : 0;
            frozenRows = rows;
            frozenColumns = cols;
        }

        var result = new
        {
            worksheetName = worksheet.Name,
            isFrozen,
            frozenRow,
            frozenColumn,
            frozenRows,
            frozenColumns,
            status = isFrozen ? "Panes are frozen" : "Panes are not frozen"
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}