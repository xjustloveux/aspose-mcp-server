using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
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
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Formula");
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
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
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

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, cell, formula, autoCalculate),
            "get" => BuildGetParameters(parameters, range),
            "get_result" => BuildGetResultParameters(parameters, cell, calculateBeforeRead),
            "calculate" => parameters,
            "set_array" => BuildSetArrayParameters(parameters, range, formula, autoCalculate),
            "get_array" => BuildCellParameters(parameters, cell),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add formula operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference.</param>
    /// <param name="formula">The formula to add.</param>
    /// <param name="autoCalculate">Whether to auto-calculate after adding.</param>
    /// <returns>OperationParameters configured for adding a formula.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? cell, string? formula,
        bool autoCalculate)
    {
        if (cell != null) parameters.Set("cell", cell);
        if (formula != null) parameters.Set("formula", formula);
        parameters.Set("autoCalculate", autoCalculate);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get formula operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range to get formulas from.</param>
    /// <returns>OperationParameters configured for getting formulas.</returns>
    private static OperationParameters BuildGetParameters(OperationParameters parameters, string? range)
    {
        if (range != null) parameters.Set("range", range);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get formula result operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference.</param>
    /// <param name="calculateBeforeRead">Whether to calculate before reading.</param>
    /// <returns>OperationParameters configured for getting formula result.</returns>
    private static OperationParameters BuildGetResultParameters(OperationParameters parameters, string? cell,
        bool calculateBeforeRead)
    {
        if (cell != null) parameters.Set("cell", cell);
        parameters.Set("calculateBeforeRead", calculateBeforeRead);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set array formula operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range for the array formula.</param>
    /// <param name="formula">The array formula.</param>
    /// <param name="autoCalculate">Whether to auto-calculate after setting.</param>
    /// <returns>OperationParameters configured for setting an array formula.</returns>
    private static OperationParameters BuildSetArrayParameters(OperationParameters parameters, string? range,
        string? formula, bool autoCalculate)
    {
        if (range != null) parameters.Set("range", range);
        if (formula != null) parameters.Set("formula", formula);
        parameters.Set("autoCalculate", autoCalculate);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters containing only the cell reference.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="cell">The cell reference.</param>
    /// <returns>OperationParameters with cell set.</returns>
    private static OperationParameters BuildCellParameters(OperationParameters parameters, string? cell)
    {
        if (cell != null) parameters.Set("cell", cell);
        return parameters;
    }
}
