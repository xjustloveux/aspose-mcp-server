using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Excel.Group;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel row and column groups (group/ungroup).
/// </summary>
[McpServerToolType]
public class ExcelGroupTool
{
    /// <summary>
    ///     Handler registry for group operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

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
        _handlerRegistry = ExcelGroupHandlerRegistry.Create();
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

        var parameters = BuildParameters(operation, sheetIndex, startRow, endRow, startColumn, endColumn, isCollapsed);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        int? startRow,
        int? endRow,
        int? startColumn,
        int? endColumn,
        bool isCollapsed)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        switch (operation.ToLowerInvariant())
        {
            case "group_rows":
                if (startRow.HasValue) parameters.Set("startRow", startRow.Value);
                if (endRow.HasValue) parameters.Set("endRow", endRow.Value);
                parameters.Set("isCollapsed", isCollapsed);
                break;

            case "ungroup_rows":
                if (startRow.HasValue) parameters.Set("startRow", startRow.Value);
                if (endRow.HasValue) parameters.Set("endRow", endRow.Value);
                break;

            case "group_columns":
                if (startColumn.HasValue) parameters.Set("startColumn", startColumn.Value);
                if (endColumn.HasValue) parameters.Set("endColumn", endColumn.Value);
                parameters.Set("isCollapsed", isCollapsed);
                break;

            case "ungroup_columns":
                if (startColumn.HasValue) parameters.Set("startColumn", startColumn.Value);
                if (endColumn.HasValue) parameters.Set("endColumn", endColumn.Value);
                break;
        }

        return parameters;
    }
}
