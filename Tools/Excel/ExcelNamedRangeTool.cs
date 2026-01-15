using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel named ranges (add, delete, get).
/// </summary>
[McpServerToolType]
public class ExcelNamedRangeTool
{
    /// <summary>
    ///     Handler registry for named range operations.
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
    ///     Initializes a new instance of the <see cref="ExcelNamedRangeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelNamedRangeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.NamedRange");
    }

    /// <summary>
    ///     Executes an Excel named range operation (add, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0). Used when range does not include sheet reference.</param>
    /// <param name="name">Name for the range. Must be a valid Excel name (required for add/delete).</param>
    /// <param name="range">Cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10', required for add).</param>
    /// <param name="comment">Comment for the named range (optional for add).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_named_range")]
    [Description(@"Manage Excel named ranges. Supports 3 operations: add, delete, get.

Usage examples:
- Add named range: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='A1:C10')
- Add with sheet reference: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='Sheet1!A1:C10')
- Delete named range: excel_named_range(operation='delete', path='book.xlsx', name='MyRange')
- Get named ranges: excel_named_range(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: add, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0). Used when range does not include sheet reference")]
        int sheetIndex = 0,
        [Description("Name for the range. Must be a valid Excel name (required for add/delete)")]
        string? name = null,
        [Description("Cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10', required for add)")]
        string? range = null,
        [Description("Comment for the named range (optional for add)")]
        string? comment = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, name, range, comment);

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

        if (operation.ToLowerInvariant() == "get")
            return result;

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
        string? name,
        string? range,
        string? comment)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        switch (operation.ToLowerInvariant())
        {
            case "add":
                if (name != null) parameters.Set("name", name);
                if (range != null) parameters.Set("range", range);
                if (comment != null) parameters.Set("comment", comment);
                break;

            case "delete":
                if (name != null) parameters.Set("name", name);
                break;

            case "get":
                break;
        }

        return parameters;
    }
}
