using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint tables (add, edit, delete, get content, insert row/column, delete row/column)
/// </summary>
[McpServerToolType]
public class PptTableTool
{
    /// <summary>
    ///     Handler registry for table operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Table");
    }

    /// <summary>
    ///     Executes a PowerPoint table operation (add, edit, delete, get_content, insert_row, insert_column, delete_row,
    ///     delete_column, edit_cell).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add, edit, delete, get_content, insert_row, insert_column,
    ///     delete_row, delete_column, edit_cell.
    /// </param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="shapeIndex">Shape index of the table (0-based, required for most operations).</param>
    /// <param name="rows">Number of rows (required for add).</param>
    /// <param name="columns">Number of columns (required for add).</param>
    /// <param name="x">X position in points (optional for add, defaults to 50).</param>
    /// <param name="y">Y position in points (optional for add, defaults to 50).</param>
    /// <param name="data">2D array of cell data as JSON (optional, for add/edit).</param>
    /// <param name="rowIndex">Row index (0-based, required for insert_row/delete_row/edit_cell).</param>
    /// <param name="columnIndex">Column index (0-based, required for insert_column/delete_column/edit_cell).</param>
    /// <param name="text">Cell text content (required for edit_cell).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_table")]
    [Description(
        @"Manage PowerPoint tables. Supports 9 operations: add, edit, delete, get_content, insert_row, insert_column, delete_row, delete_column, edit_cell.

Coordinate unit: 1 inch = 72 points.

Usage examples:
- Add table: ppt_table(operation='add', path='presentation.pptx', slideIndex=0, rows=3, columns=3, x=100, y=100)
- Edit table: ppt_table(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, data='[[""A"",""B""],[""C"",""D""]]')
- Delete table: ppt_table(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get content: ppt_table(operation='get_content', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Insert row: ppt_table(operation='insert_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Insert column: ppt_table(operation='insert_column', path='presentation.pptx', slideIndex=0, shapeIndex=0, columnIndex=1)
- Delete row: ppt_table(operation='delete_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Edit cell: ppt_table(operation='edit_cell', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=0, columnIndex=0, text='New Value')")]
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
        [Description(@"Operation to perform.
- 'add': Add a table (required params: path, slideIndex, rows, columns)
- 'edit': Edit table data (required params: path, slideIndex, shapeIndex, data)
- 'delete': Delete a table (required params: path, slideIndex, shapeIndex)
- 'get_content': Get table content (required params: path, slideIndex, shapeIndex)
- 'insert_row': Insert a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'insert_column': Insert a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'delete_row': Delete a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'delete_column': Delete a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'edit_cell': Edit cell content (required params: path, slideIndex, shapeIndex, rowIndex, columnIndex, text)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int slideIndex = 0,
        [Description("Shape index of the table (0-based, required for most operations)")]
        int? shapeIndex = null,
        [Description("Number of rows (required for add)")]
        int? rows = null,
        [Description("Number of columns (required for add)")]
        int? columns = null,
        [Description("X position in points (optional for add, defaults to 50)")]
        float x = 50,
        [Description("Y position in points (optional for add, defaults to 50)")]
        float y = 50,
        [Description("2D array of cell data as JSON (optional, for add/edit)")]
        string? data = null,
        [Description("Row index (0-based, required for insert_row/delete_row/edit_cell)")]
        int? rowIndex = null,
        [Description("Column index (0-based, required for insert_column/delete_column/edit_cell)")]
        int? columnIndex = null,
        [Description("Cell text content (required for edit_cell)")]
        string? text = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, rows, columns, x, y, data,
            rowIndex, columnIndex, text);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get_content", StringComparison.OrdinalIgnoreCase))
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
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        int slideIndex,
        int? shapeIndex,
        int? rows,
        int? columns,
        float x,
        float y,
        string? data,
        int? rowIndex,
        int? columnIndex,
        string? text)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, rows, columns, x, y, data),
            "edit" => BuildEditParameters(parameters, shapeIndex, data),
            "delete" or "get_content" => BuildShapeIndexParameters(parameters, shapeIndex),
            "insert_row" or "delete_row" => BuildRowParameters(parameters, shapeIndex, rowIndex),
            "insert_column" or "delete_column" => BuildColumnParameters(parameters, shapeIndex, columnIndex),
            "edit_cell" => BuildEditCellParameters(parameters, shapeIndex, rowIndex, columnIndex, text),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add table operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="rows">The number of rows for the table.</param>
    /// <param name="columns">The number of columns for the table.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="data">The 2D array of cell data as JSON.</param>
    /// <returns>OperationParameters configured for the add operation.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, int? rows, int? columns,
        float x, float y, string? data)
    {
        if (rows.HasValue) parameters.Set("rows", rows.Value);
        if (columns.HasValue) parameters.Set("columns", columns.Value);
        parameters.Set("x", x);
        parameters.Set("y", y);
        if (data != null) parameters.Set("data", data);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit table operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="data">The 2D array of cell data as JSON.</param>
    /// <returns>OperationParameters configured for the edit operation.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int? shapeIndex,
        string? data)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (data != null) parameters.Set("data", data);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for shape index-based operations (delete, get_content).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <returns>OperationParameters configured for shape index-based operations.</returns>
    private static OperationParameters BuildShapeIndexParameters(OperationParameters parameters, int? shapeIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for row operations (insert_row, delete_row).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="rowIndex">The row index (0-based).</param>
    /// <returns>OperationParameters configured for row operations.</returns>
    private static OperationParameters BuildRowParameters(OperationParameters parameters, int? shapeIndex,
        int? rowIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (rowIndex.HasValue) parameters.Set("rowIndex", rowIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for column operations (insert_column, delete_column).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="columnIndex">The column index (0-based).</param>
    /// <returns>OperationParameters configured for column operations.</returns>
    private static OperationParameters BuildColumnParameters(OperationParameters parameters, int? shapeIndex,
        int? columnIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (columnIndex.HasValue) parameters.Set("columnIndex", columnIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit cell operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="rowIndex">The row index (0-based).</param>
    /// <param name="columnIndex">The column index (0-based).</param>
    /// <param name="text">The cell text content.</param>
    /// <returns>OperationParameters configured for the edit cell operation.</returns>
    private static OperationParameters BuildEditCellParameters(OperationParameters parameters, int? shapeIndex,
        int? rowIndex, int? columnIndex, string? text)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (rowIndex.HasValue) parameters.Set("rowIndex", rowIndex.Value);
        if (columnIndex.HasValue) parameters.Set("columnIndex", columnIndex.Value);
        if (text != null) parameters.Set("text", text);
        return parameters;
    }
}
