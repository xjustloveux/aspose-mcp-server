using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel shapes (add, edit, delete, get, add_textbox).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Shape")]
[McpServerToolType]
public class ExcelShapeTool
{
    /// <summary>
    ///     Handler registry for shape operations.
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
    ///     Initializes a new instance of the <see cref="ExcelShapeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelShapeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Shape");
    }

    /// <summary>
    ///     Executes an Excel shape operation (add, edit, delete, get, add_textbox).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="shapeIndex">Shape index (0-based, for edit/delete).</param>
    /// <param name="shapeType">Auto shape type (for add, e.g., 'Rectangle', 'Oval').</param>
    /// <param name="text">Text content (for add/add_textbox/edit).</param>
    /// <param name="name">Shape name (for edit).</param>
    /// <param name="upperLeftRow">Upper-left row (for add/add_textbox/edit).</param>
    /// <param name="upperLeftColumn">Upper-left column (for add/add_textbox/edit).</param>
    /// <param name="width">Width in pixels (for add/add_textbox/edit).</param>
    /// <param name="height">Height in pixels (for add/add_textbox/edit).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_shape",
        Title = "Excel Shape Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel shapes. Supports 5 operations: add, edit, delete, get, add_textbox.

Usage examples:
- Add shape: excel_shape(operation='add', path='book.xlsx', shapeType='Rectangle', upperLeftRow=1, upperLeftColumn=1)
- Add textbox: excel_shape(operation='add_textbox', path='book.xlsx', text='Hello World')
- Get shapes: excel_shape(operation='get', path='book.xlsx')
- Edit shape: excel_shape(operation='edit', path='book.xlsx', shapeIndex=0, text='Updated')
- Delete shape: excel_shape(operation='delete', path='book.xlsx', shapeIndex=0)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add an auto shape (required params: shapeType)
- 'get': Get shapes information (optional: shapeIndex)
- 'edit': Edit shape properties (required params: shapeIndex)
- 'delete': Delete a shape (required params: shapeIndex)
- 'add_textbox': Add a textbox (required params: text)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Shape index (0-based, for edit/delete)")]
        int? shapeIndex = null,
        [Description("Auto shape type (for add, e.g., 'Rectangle', 'Oval', 'Star5')")]
        string? shapeType = null,
        [Description("Text content (for add/add_textbox/edit)")]
        string? text = null,
        [Description("Shape name (for edit)")] string? name = null,
        [Description("Upper-left row (for add/add_textbox/edit)")]
        int? upperLeftRow = null,
        [Description("Upper-left column (for add/add_textbox/edit)")]
        int? upperLeftColumn = null,
        [Description("Width in pixels (for add/add_textbox/edit)")]
        int? width = null,
        [Description("Height in pixels (for add/add_textbox/edit)")]
        int? height = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, shapeIndex, shapeType, text, name,
            upperLeftRow, upperLeftColumn, width, height);

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
        int? shapeIndex,
        string? shapeType,
        string? text,
        string? name,
        int? upperLeftRow,
        int? upperLeftColumn,
        int? width,
        int? height)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, shapeType, text, upperLeftRow, upperLeftColumn, width, height),
            "add_textbox" => BuildAddTextBoxParameters(parameters, text, upperLeftRow, upperLeftColumn, width, height),
            "get" => BuildGetParameters(parameters, shapeIndex),
            "edit" => BuildEditParameters(parameters, shapeIndex, text, name, width, height, upperLeftRow,
                upperLeftColumn),
            "delete" => BuildDeleteParameters(parameters, shapeIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add operation.
    /// </summary>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? shapeType,
        string? text, int? upperLeftRow, int? upperLeftColumn, int? width, int? height)
    {
        if (shapeType != null) parameters.Set("shapeType", shapeType);
        if (text != null) parameters.Set("text", text);
        if (upperLeftRow.HasValue) parameters.Set("upperLeftRow", upperLeftRow.Value);
        if (upperLeftColumn.HasValue) parameters.Set("upperLeftColumn", upperLeftColumn.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_textbox operation.
    /// </summary>
    private static OperationParameters BuildAddTextBoxParameters(OperationParameters parameters, string? text,
        int? upperLeftRow, int? upperLeftColumn, int? width, int? height)
    {
        if (text != null) parameters.Set("text", text);
        if (upperLeftRow.HasValue) parameters.Set("upperLeftRow", upperLeftRow.Value);
        if (upperLeftColumn.HasValue) parameters.Set("upperLeftColumn", upperLeftColumn.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get operation.
    /// </summary>
    private static OperationParameters BuildGetParameters(OperationParameters parameters, int? shapeIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit operation.
    /// </summary>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int? shapeIndex,
        string? text, string? name, int? width, int? height, int? upperLeftRow, int? upperLeftColumn)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (text != null) parameters.Set("text", text);
        if (name != null) parameters.Set("name", name);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (upperLeftRow.HasValue) parameters.Set("upperLeftRow", upperLeftRow.Value);
        if (upperLeftColumn.HasValue) parameters.Set("upperLeftColumn", upperLeftColumn.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? shapeIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }
}
