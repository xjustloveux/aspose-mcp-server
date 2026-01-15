using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel cells (write, edit, get, clear)
/// </summary>
[McpServerToolType]
public class ExcelCellTool
{
    /// <summary>
    ///     Handler registry for cell operations.
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
    ///     Initializes a new instance of the <see cref="ExcelCellTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelCellTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Cell");
    }

    /// <summary>
    ///     Executes an Excel cell operation (write, edit, get, or clear).
    /// </summary>
    /// <param name="operation">The operation to perform: write, edit, get, or clear.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="cell">Cell reference (e.g., 'A1', 'B2', 'AA100').</param>
    /// <param name="value">Value to write.</param>
    /// <param name="formula">Formula to set (optional, for edit, overrides value).</param>
    /// <param name="clearValue">Clear cell value (optional, for edit).</param>
    /// <param name="calculateFormula">Calculate formulas before reading value (optional, for get).</param>
    /// <param name="includeFormula">Include formula if present (optional, for get).</param>
    /// <param name="includeFormat">Include format information (optional, for get).</param>
    /// <param name="clearContent">Clear cell content (optional, for clear).</param>
    /// <param name="clearFormat">Clear cell format (optional, for clear).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when cell is not provided or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_cell")]
    [Description(@"Manage Excel cells. Supports 4 operations: write, edit, get, clear.

Usage examples:
- Write cell: excel_cell(operation='write', path='book.xlsx', cell='A1', value='Hello')
- Edit cell: excel_cell(operation='edit', path='book.xlsx', cell='A1', value='Updated')
- Get cell: excel_cell(operation='get', path='book.xlsx', cell='A1')
- Clear cell: excel_cell(operation='clear', path='book.xlsx', cell='A1')")]
    public string Execute(
        [Description("Operation: write, edit, get, clear")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', 'B2', 'AA100')")]
        string? cell = null,
        [Description("Value to write")] string? value = null,
        [Description("Formula to set (optional, for edit, overrides value)")]
        string? formula = null,
        [Description("Clear cell value (optional, for edit)")]
        bool clearValue = false,
        [Description("Calculate formulas before reading value (optional, for get, default: false)")]
        bool calculateFormula = false,
        [Description("Include formula if present (optional, for get, default: true)")]
        bool includeFormula = true,
        [Description("Include format information (optional, for get, default: false)")]
        bool includeFormat = false,
        [Description("Clear cell content (optional, for clear, default: true)")]
        bool clearContent = true,
        [Description("Clear cell format (optional, for clear, default: false)")]
        bool clearFormat = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, cell, value, formula, clearValue,
            calculateFormula, includeFormula, includeFormat, clearContent, clearFormat);

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
        string? cell,
        string? value,
        string? formula,
        bool clearValue,
        bool calculateFormula,
        bool includeFormula,
        bool includeFormat,
        bool clearContent,
        bool clearFormat)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        if (cell != null) parameters.Set("cell", cell);

        switch (operation.ToLowerInvariant())
        {
            case "write":
                if (value != null) parameters.Set("value", value);
                break;

            case "edit":
                if (value != null) parameters.Set("value", value);
                if (formula != null) parameters.Set("formula", formula);
                parameters.Set("clearValue", clearValue);
                break;

            case "get":
                parameters.Set("calculateFormula", calculateFormula);
                parameters.Set("includeFormula", includeFormula);
                parameters.Set("includeFormat", includeFormat);
                break;

            case "clear":
                parameters.Set("clearContent", clearContent);
                parameters.Set("clearFormat", clearFormat);
                break;
        }

        return parameters;
    }
}
