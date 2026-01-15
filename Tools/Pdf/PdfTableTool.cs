using System.ComponentModel;
using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Unified tool for managing PDF tables (add, edit)
/// </summary>
[McpServerToolType]
public class PdfTableTool
{
    /// <summary>
    ///     Handler registry for table operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Pdf.Table");
    }

    /// <summary>
    ///     Executes a PDF table operation (add, edit).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="pageIndex">Page index (1-based, required for add).</param>
    /// <param name="rows">Number of rows (required for add).</param>
    /// <param name="columns">Number of columns (required for add).</param>
    /// <param name="data">Table data (array of arrays, for add).</param>
    /// <param name="x">X position (left margin) in PDF points (for add).</param>
    /// <param name="y">Y position (top margin) in PDF points (for add).</param>
    /// <param name="columnWidths">Space-separated column widths in PDF points (for add).</param>
    /// <param name="tableIndex">Table index (0-based, required for edit).</param>
    /// <param name="cellRow">Cell row index (0-based, for edit).</param>
    /// <param name="cellColumn">Cell column index (0-based, for edit).</param>
    /// <param name="cellValue">New cell value (for edit).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_table")]
    [Description(@"Manage tables in PDF documents. Supports 2 operations: add, edit.

Usage examples:
- Add table: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=3, columns=3, data=[['A','B','C'],['1','2','3']])
- Add table with position: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=2, x=100, y=500)
- Add table with column widths: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=3, columnWidths='100 150 200')
- Edit table cell: pdf_table(operation='edit', path='doc.pdf', tableIndex=0, cellRow=0, cellColumn=1, cellValue='NewValue')

Note: PDF table editing has limitations. After saving, tables may be converted to graphics and cannot be edited as Table objects.")]
    public string Execute(
        [Description("Operation: add, edit")] string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add)")]
        int pageIndex = 1,
        [Description("Number of rows (required for add)")]
        int rows = 0,
        [Description("Number of columns (required for add)")]
        int columns = 0,
        [Description("Table data (array of arrays, for add)")]
        string[][]? data = null,
        [Description("X position (left margin) in PDF points (for add, default: 100)")]
        double x = 100,
        [Description("Y position (top margin) in PDF points (for add, default: 600)")]
        double y = 600,
        [Description("Space-separated column widths in PDF points (for add, e.g., '100 150 200')")]
        string? columnWidths = null,
        [Description("Table index (0-based, required for edit)")]
        int tableIndex = 0,
        [Description("Cell row index (0-based, for edit)")]
        int? cellRow = null,
        [Description("Cell column index (0-based, for edit)")]
        int? cellColumn = null,
        [Description("New cell value (for edit)")]
        string? cellValue = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, pageIndex, rows, columns, data, x, y, columnWidths,
            tableIndex, cellRow, cellColumn, cellValue);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
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
        int pageIndex,
        int rows,
        int columns,
        string[][]? data,
        double x,
        double y,
        string? columnWidths,
        int tableIndex,
        int? cellRow,
        int? cellColumn,
        string? cellValue)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                parameters.Set("pageIndex", pageIndex);
                parameters.Set("rows", rows);
                parameters.Set("columns", columns);
                if (data != null) parameters.Set("data", data);
                parameters.Set("x", x);
                parameters.Set("y", y);
                if (columnWidths != null) parameters.Set("columnWidths", columnWidths);
                break;

            case "edit":
                parameters.Set("tableIndex", tableIndex);
                if (cellRow.HasValue) parameters.Set("cellRow", cellRow.Value);
                if (cellColumn.HasValue) parameters.Set("cellColumn", cellColumn.Value);
                if (cellValue != null) parameters.Set("cellValue", cellValue);
                break;
        }

        return parameters;
    }
}
