using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
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
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Group");
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
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
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
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107
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

        return operation.ToLowerInvariant() switch
        {
            "group_rows" => BuildGroupRowsParameters(parameters, startRow, endRow, isCollapsed),
            "ungroup_rows" => BuildUngroupRowsParameters(parameters, startRow, endRow),
            "group_columns" => BuildGroupColumnsParameters(parameters, startColumn, endColumn, isCollapsed),
            "ungroup_columns" => BuildUngroupColumnsParameters(parameters, startColumn, endColumn),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the group rows operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="startRow">The start row index.</param>
    /// <param name="endRow">The end row index.</param>
    /// <param name="isCollapsed">Whether to collapse the group initially.</param>
    /// <returns>OperationParameters configured for grouping rows.</returns>
    private static OperationParameters BuildGroupRowsParameters(OperationParameters parameters, int? startRow,
        int? endRow, bool isCollapsed)
    {
        if (startRow.HasValue) parameters.Set("startRow", startRow.Value);
        if (endRow.HasValue) parameters.Set("endRow", endRow.Value);
        parameters.Set("isCollapsed", isCollapsed);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the ungroup rows operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="startRow">The start row index.</param>
    /// <param name="endRow">The end row index.</param>
    /// <returns>OperationParameters configured for ungrouping rows.</returns>
    private static OperationParameters BuildUngroupRowsParameters(OperationParameters parameters, int? startRow,
        int? endRow)
    {
        if (startRow.HasValue) parameters.Set("startRow", startRow.Value);
        if (endRow.HasValue) parameters.Set("endRow", endRow.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the group columns operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="startColumn">The start column index.</param>
    /// <param name="endColumn">The end column index.</param>
    /// <param name="isCollapsed">Whether to collapse the group initially.</param>
    /// <returns>OperationParameters configured for grouping columns.</returns>
    private static OperationParameters BuildGroupColumnsParameters(OperationParameters parameters, int? startColumn,
        int? endColumn, bool isCollapsed)
    {
        if (startColumn.HasValue) parameters.Set("startColumn", startColumn.Value);
        if (endColumn.HasValue) parameters.Set("endColumn", endColumn.Value);
        parameters.Set("isCollapsed", isCollapsed);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the ungroup columns operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="startColumn">The start column index.</param>
    /// <param name="endColumn">The end column index.</param>
    /// <returns>OperationParameters configured for ungrouping columns.</returns>
    private static OperationParameters BuildUngroupColumnsParameters(OperationParameters parameters, int? startColumn,
        int? endColumn)
    {
        if (startColumn.HasValue) parameters.Set("startColumn", startColumn.Value);
        if (endColumn.HasValue) parameters.Set("endColumn", endColumn.Value);
        return parameters;
    }
}
