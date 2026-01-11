using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Excel.Formula;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel formulas (add, get, get_result, calculate, set_array, get_array).
/// </summary>
[McpServerToolType]
public class ExcelFormulaTool
{
    /// <summary>
    ///     Handler registry for formula operations.
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
    ///     Initializes a new instance of the <see cref="ExcelFormulaTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelFormulaTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = ExcelFormulaHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes an Excel formula operation (add, get, get_result, calculate, set_array, get_array).
    /// </summary>
    /// <param name="operation">The operation to perform: add, get, get_result, calculate, set_array, get_array.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="cell">Cell reference (e.g., 'A1', required for add/get_result/get_array).</param>
    /// <param name="range">Cell range (e.g., 'A1:C10', optional for get, required for set_array).</param>
    /// <param name="formula">Formula (e.g., '=SUM(A1:A10)', required for add/set_array).</param>
    /// <param name="calculateBeforeRead">Calculate formulas before reading (optional, for get_result, default: true).</param>
    /// <param name="autoCalculate">Automatically calculate formulas after adding (optional, for add/set_array, default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_formula")]
    [Description(@"Manage Excel formulas. Supports 6 operations: add, get, get_result, calculate, set_array, get_array.

Usage examples:
- Add formula: excel_formula(operation='add', path='book.xlsx', cell='A1', formula='=SUM(B1:B10)')
- Get formula: excel_formula(operation='get', path='book.xlsx', cell='A1')
- Get result: excel_formula(operation='get_result', path='book.xlsx', cell='A1')
- Calculate: excel_formula(operation='calculate', path='book.xlsx')
- Set array formula: excel_formula(operation='set_array', path='book.xlsx', range='A1:A10', formula='=B1:B10*2')
- Get array formula: excel_formula(operation='get_array', path='book.xlsx', cell='A1')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a formula to a cell (required params: path, cell, formula)
- 'get': Get formula from a cell (required params: path, cell)
- 'get_result': Get formula result (required params: path, cell)
- 'calculate': Calculate all formulas (required params: path)
- 'set_array': Set array formula (required params: path, range, formula)
- 'get_array': Get array formula (required params: path, cell)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', required for add/get_result/get_array)")]
        string? cell = null,
        [Description("Cell range (e.g., 'A1:C10', optional for get, required for set_array)")]
        string? range = null,
        [Description("Formula (e.g., '=SUM(A1:A10)', required for add/set_array)")]
        string? formula = null,
        [Description("Calculate formulas before reading (optional, for get_result, default: true)")]
        bool calculateBeforeRead = true,
        [Description("Automatically calculate formulas after adding (optional, for add/set_array, default: true)")]
        bool autoCalculate = true)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters =
            BuildParameters(operation, sheetIndex, cell, range, formula, calculateBeforeRead, autoCalculate);

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

        var op = operation.ToLowerInvariant();
        if (op == "get" || op == "get_result" || op == "get_array")
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
        string? range,
        string? formula,
        bool calculateBeforeRead,
        bool autoCalculate)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        switch (operation.ToLowerInvariant())
        {
            case "add":
                if (cell != null) parameters.Set("cell", cell);
                if (formula != null) parameters.Set("formula", formula);
                parameters.Set("autoCalculate", autoCalculate);
                break;

            case "get":
                if (range != null) parameters.Set("range", range);
                break;

            case "get_result":
                if (cell != null) parameters.Set("cell", cell);
                parameters.Set("calculateBeforeRead", calculateBeforeRead);
                break;

            case "calculate":
                break;

            case "set_array":
                if (range != null) parameters.Set("range", range);
                if (formula != null) parameters.Set("formula", formula);
                parameters.Set("autoCalculate", autoCalculate);
                break;

            case "get_array":
                if (cell != null) parameters.Set("cell", cell);
                break;
        }

        return parameters;
    }
}
