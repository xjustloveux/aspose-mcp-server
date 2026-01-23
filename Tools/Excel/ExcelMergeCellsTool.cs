using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel merged cells (merge, unmerge, get).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.MergeCells")]
[McpServerToolType]
public class ExcelMergeCellsTool
{
    /// <summary>
    ///     Handler registry for merge cells operations.
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
    ///     Initializes a new instance of the <see cref="ExcelMergeCellsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelMergeCellsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.MergeCells");
    }

    /// <summary>
    ///     Executes an Excel merge cells operation (merge, unmerge, get).
    /// </summary>
    /// <param name="operation">The operation to perform: merge, unmerge, get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="range">
    ///     Cell range to merge/unmerge (e.g., 'A1:C3', must include at least 2 cells, required for
    ///     merge/unmerge).
    /// </param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_merge_cells",
        Title = "Excel Merge Cells Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel merged cells. Supports 3 operations: merge, unmerge, get.

Usage examples:
- Merge cells: excel_merge_cells(operation='merge', path='book.xlsx', range='A1:C1')
- Unmerge cells: excel_merge_cells(operation='unmerge', path='book.xlsx', range='A1:C1')
- Get merged cells: excel_merge_cells(operation='get', path='book.xlsx')

WARNING: Merging cells will only keep the value of the top-left cell. All other cell values will be lost.")]
    public object Execute(
        [Description("Operation: merge, unmerge, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Cell range to merge/unmerge (e.g., 'A1:C3', must include at least 2 cells, required for merge/unmerge)")]
        string? range = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range);

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

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? range)
    {
        return operation.ToLowerInvariant() switch
        {
            "merge" or "unmerge" => BuildMergeUnmergeParameters(sheetIndex, range),
            "get" => BuildGetParameters(sheetIndex),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the merge or unmerge cells operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="range">The cell range to merge/unmerge (e.g., 'A1:C3').</param>
    /// <returns>OperationParameters configured for merge/unmerge operation.</returns>
    private static OperationParameters BuildMergeUnmergeParameters(int sheetIndex, string? range)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        if (range != null) parameters.Set("range", range);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get merged cells operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <returns>OperationParameters configured for getting merged cells.</returns>
    private static OperationParameters BuildGetParameters(int sheetIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        return parameters;
    }
}
