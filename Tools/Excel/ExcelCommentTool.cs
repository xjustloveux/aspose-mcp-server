using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel comments (add, edit, delete, get)
/// </summary>
[McpServerToolType]
public class ExcelCommentTool
{
    /// <summary>
    ///     Handler registry for comment operations.
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
    ///     Initializes a new instance of the <see cref="ExcelCommentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelCommentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Comment");
    }

    /// <summary>
    ///     Executes an Excel comment operation (add, edit, delete, or get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, or get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="cell">Cell reference (e.g., 'A1', required for add/edit/delete, optional for get).</param>
    /// <param name="comment">Comment text (required for add/edit).</param>
    /// <param name="author">Comment author (optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_comment")]
    [Description(@"Manage Excel comments. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add comment: excel_comment(operation='add', path='book.xlsx', cell='A1', comment='This is a comment')
- Edit comment: excel_comment(operation='edit', path='book.xlsx', cell='A1', comment='Updated comment')
- Delete comment: excel_comment(operation='delete', path='book.xlsx', cell='A1')
- Get comments: excel_comment(operation='get', path='book.xlsx')")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', required for add/edit/delete, optional for get)")]
        string? cell = null,
        [Description("Comment text (required for add/edit)")]
        string? comment = null,
        [Description("Comment author (optional)")]
        string? author = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, cell, comment, author);

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
            return result;

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
        string? cell,
        string? comment,
        string? author)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" or "edit" => BuildAddEditParameters(parameters, cell, comment, author),
            "delete" or "get" => BuildCellParameters(parameters, cell),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add or edit comment operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference for the comment.</param>
    /// <param name="comment">The comment text.</param>
    /// <param name="author">The comment author.</param>
    /// <returns>OperationParameters configured for adding or editing comment.</returns>
    private static OperationParameters BuildAddEditParameters(OperationParameters parameters, string? cell,
        string? comment, string? author)
    {
        if (cell != null) parameters.Set("cell", cell);
        if (comment != null) parameters.Set("comment", comment);
        if (author != null) parameters.Set("author", author);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for operations that require only cell reference.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference.</param>
    /// <returns>OperationParameters configured with cell reference.</returns>
    private static OperationParameters BuildCellParameters(OperationParameters parameters, string? cell)
    {
        if (cell != null) parameters.Set("cell", cell);
        return parameters;
    }
}
