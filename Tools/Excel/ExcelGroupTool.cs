using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel row and column groups (group/ungroup).
/// </summary>
[McpServerToolType]
public class ExcelGroupTool
{
    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelGroupTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelGroupTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel group operation (group_rows, ungroup_rows, group_columns, ungroup_columns).
    /// </summary>
    /// <param name="operation">The operation to perform: group_rows, ungroup_rows, group_columns, ungroup_columns.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="startRow">Start row index (0-based, required for group_rows/ungroup_rows).</param>
    /// <param name="endRow">End row index (0-based, must be >= startRow, required for group_rows/ungroup_rows).</param>
    /// <param name="startColumn">Start column index (0-based, required for group_columns/ungroup_columns).</param>
    /// <param name="endColumn">End column index (0-based, must be >= startColumn, required for group_columns/ungroup_columns).</param>
    /// <param name="isCollapsed">Collapse group initially (optional, for group_rows/group_columns, default: false).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_group")]
    [Description(@"Manage Excel groups. Supports 4 operations: group_rows, ungroup_rows, group_columns, ungroup_columns.

Usage examples:
- Group rows: excel_group(operation='group_rows', path='book.xlsx', startRow=1, endRow=5)
- Ungroup rows: excel_group(operation='ungroup_rows', path='book.xlsx', startRow=1, endRow=5)
- Group columns: excel_group(operation='group_columns', path='book.xlsx', startColumn=1, endColumn=3)
- Ungroup columns: excel_group(operation='ungroup_columns', path='book.xlsx', startColumn=1, endColumn=3)")]
    public string Execute(
        [Description("Operation: group_rows, ungroup_rows, group_columns, ungroup_columns")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Start row index (0-based, required for group_rows/ungroup_rows)")]
        int? startRow = null,
        [Description("End row index (0-based, must be >= startRow, required for group_rows/ungroup_rows)")]
        int? endRow = null,
        [Description("Start column index (0-based, required for group_columns/ungroup_columns)")]
        int? startColumn = null,
        [Description("End column index (0-based, must be >= startColumn, required for group_columns/ungroup_columns)")]
        int? endColumn = null,
        [Description("Collapse group initially (optional, for group_rows/group_columns, default: false)")]
        bool isCollapsed = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "group_rows" => GroupRows(ctx, outputPath, sheetIndex, startRow, endRow, isCollapsed),
            "ungroup_rows" => UngroupRows(ctx, outputPath, sheetIndex, startRow, endRow),
            "group_columns" => GroupColumns(ctx, outputPath, sheetIndex, startColumn, endColumn, isCollapsed),
            "ungroup_columns" => UngroupColumns(ctx, outputPath, sheetIndex, startColumn, endColumn),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Groups rows together.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="startRow">The start row index (0-based).</param>
    /// <param name="endRow">The end row index (0-based).</param>
    /// <param name="isCollapsed">Whether to collapse the group initially.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when startRow or endRow is not provided, or when the range is invalid.</exception>
    private static string GroupRows(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int? startRow,
        int? endRow, bool isCollapsed)
    {
        if (!startRow.HasValue)
            throw new ArgumentException("Operation 'group_rows' requires parameter 'startRow'.");
        if (!endRow.HasValue)
            throw new ArgumentException("Operation 'group_rows' requires parameter 'endRow'.");

        ValidateRowRange(startRow.Value, endRow.Value);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupRows(startRow.Value, endRow.Value, isCollapsed);

        ctx.Save(outputPath);
        return $"Rows {startRow}-{endRow} grouped in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Ungroups rows.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="startRow">The start row index (0-based).</param>
    /// <param name="endRow">The end row index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when startRow or endRow is not provided, or when the range is invalid.</exception>
    private static string UngroupRows(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int? startRow,
        int? endRow)
    {
        if (!startRow.HasValue)
            throw new ArgumentException("Operation 'ungroup_rows' requires parameter 'startRow'.");
        if (!endRow.HasValue)
            throw new ArgumentException("Operation 'ungroup_rows' requires parameter 'endRow'.");

        ValidateRowRange(startRow.Value, endRow.Value);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupRows(startRow.Value, endRow.Value);

        ctx.Save(outputPath);
        return $"Rows {startRow}-{endRow} ungrouped in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Groups columns together.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="startColumn">The start column index (0-based).</param>
    /// <param name="endColumn">The end column index (0-based).</param>
    /// <param name="isCollapsed">Whether to collapse the group initially.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when startColumn or endColumn is not provided, or when the range is invalid.</exception>
    private static string GroupColumns(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? startColumn, int? endColumn, bool isCollapsed)
    {
        if (!startColumn.HasValue)
            throw new ArgumentException("Operation 'group_columns' requires parameter 'startColumn'.");
        if (!endColumn.HasValue)
            throw new ArgumentException("Operation 'group_columns' requires parameter 'endColumn'.");

        ValidateColumnRange(startColumn.Value, endColumn.Value);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.GroupColumns(startColumn.Value, endColumn.Value, isCollapsed);

        ctx.Save(outputPath);
        return $"Columns {startColumn}-{endColumn} grouped in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Ungroups columns.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="startColumn">The start column index (0-based).</param>
    /// <param name="endColumn">The end column index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when startColumn or endColumn is not provided, or when the range is invalid.</exception>
    private static string UngroupColumns(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? startColumn, int? endColumn)
    {
        if (!startColumn.HasValue)
            throw new ArgumentException("Operation 'ungroup_columns' requires parameter 'startColumn'.");
        if (!endColumn.HasValue)
            throw new ArgumentException("Operation 'ungroup_columns' requires parameter 'endColumn'.");

        ValidateColumnRange(startColumn.Value, endColumn.Value);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Cells.UngroupColumns(startColumn.Value, endColumn.Value);

        ctx.Save(outputPath);
        return $"Columns {startColumn}-{endColumn} ungrouped in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Validates row range indices.
    /// </summary>
    /// <param name="startRow">The start row index (0-based).</param>
    /// <param name="endRow">The end row index (0-based).</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when startRow or endRow is negative, or when startRow is greater than
    ///     endRow.
    /// </exception>
    private static void ValidateRowRange(int startRow, int endRow)
    {
        if (startRow < 0)
            throw new ArgumentException($"startRow cannot be negative. Got: {startRow}");
        if (endRow < 0)
            throw new ArgumentException($"endRow cannot be negative. Got: {endRow}");
        if (startRow > endRow)
            throw new ArgumentException($"startRow ({startRow}) cannot be greater than endRow ({endRow}).");
    }

    /// <summary>
    ///     Validates column range indices.
    /// </summary>
    /// <param name="startColumn">The start column index (0-based).</param>
    /// <param name="endColumn">The end column index (0-based).</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when startColumn or endColumn is negative, or when startColumn is greater
    ///     than endColumn.
    /// </exception>
    private static void ValidateColumnRange(int startColumn, int endColumn)
    {
        if (startColumn < 0)
            throw new ArgumentException($"startColumn cannot be negative. Got: {startColumn}");
        if (endColumn < 0)
            throw new ArgumentException($"endColumn cannot be negative. Got: {endColumn}");
        if (startColumn > endColumn)
            throw new ArgumentException($"startColumn ({startColumn}) cannot be greater than endColumn ({endColumn}).");
    }
}