using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel conditional formatting (add, edit, delete, get)
/// </summary>
[McpServerToolType]
public class ExcelConditionalFormattingTool
{
    /// <summary>
    ///     Handler registry for conditional formatting operations.
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
    ///     Initializes a new instance of the <see cref="ExcelConditionalFormattingTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelConditionalFormattingTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.ConditionalFormatting");
    }

    /// <summary>
    ///     Executes an Excel conditional formatting operation (add, edit, delete, or get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, or get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range (e.g., 'A1:A10', required for add).</param>
    /// <param name="conditionalFormattingIndex">Conditional formatting index (0-based, required for edit/delete).</param>
    /// <param name="conditionIndex">Condition index within the formatting rule (0-based, optional for edit).</param>
    /// <param name="condition">Condition type: GreaterThan, LessThan, Between, Equal (required for add).</param>
    /// <param name="value">Condition value / Formula1 (required for add).</param>
    /// <param name="formula2">Second value for 'Between' condition (optional).</param>
    /// <param name="backgroundColor">Background color for matching cells.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_conditional_formatting")]
    [Description(@"Manage Excel conditional formatting. Supports 4 operations: add, edit, delete, get.

You can add multiple conditional formatting rules to the same range by calling the 'add' operation multiple times. Each rule is independent and will be evaluated separately. To add multiple rules, simply call the 'add' operation multiple times with different conditions for the same range.

Usage examples:
- Add conditional formatting: excel_conditional_formatting(operation='add', path='book.xlsx', range='A1:A10', condition='Between', value='10', formula2='100', backgroundColor='#FF0000')
- Add multiple rules: Call 'add' multiple times with different conditions to create multiple rules for the same range
- Edit conditional formatting: excel_conditional_formatting(operation='edit', path='book.xlsx', conditionalFormattingIndex=0, condition='GreaterThan', value='50')
- Delete conditional formatting: excel_conditional_formatting(operation='delete', path='book.xlsx', conditionalFormattingIndex=0)
- Get conditional formatting: excel_conditional_formatting(operation='get', path='book.xlsx', range='A1:A10')")]
    public string Execute(
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
        [Description("Cell range (e.g., 'A1:A10', required for add)")]
        string? range = null,
        [Description("Conditional formatting index (0-based, required for edit/delete)")]
        int conditionalFormattingIndex = 0,
        [Description("Condition index within the formatting rule (0-based, optional for edit)")]
        int? conditionIndex = null,
        [Description("Condition type: GreaterThan, LessThan, Between, Equal (required for add)")]
        string? condition = null,
        [Description("Condition value / Formula1 (required for add)")]
        string? value = null,
        [Description("Second value for 'Between' condition (optional)")]
        string? formula2 = null,
        [Description("Background color for matching cells (default: Yellow)")]
        string backgroundColor = "Yellow")
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, conditionalFormattingIndex,
            conditionIndex, condition, value, formula2, backgroundColor);

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
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? range,
        int conditionalFormattingIndex,
        int? conditionIndex,
        string? condition,
        string? value,
        string? formula2,
        string backgroundColor)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, range, condition, value, formula2, backgroundColor),
            "edit" => BuildEditParameters(parameters, conditionalFormattingIndex, conditionIndex, condition, value,
                formula2, backgroundColor),
            "delete" => BuildDeleteParameters(parameters, conditionalFormattingIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add conditional formatting operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range to apply formatting.</param>
    /// <param name="condition">The condition type.</param>
    /// <param name="value">The condition value or first formula.</param>
    /// <param name="formula2">The second value for between condition.</param>
    /// <param name="backgroundColor">The background color for matching cells.</param>
    /// <returns>OperationParameters configured for adding conditional formatting.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? range,
        string? condition, string? value, string? formula2, string backgroundColor)
    {
        if (range != null) parameters.Set("range", range);
        if (condition != null) parameters.Set("condition", condition);
        if (value != null) parameters.Set("value", value);
        if (formula2 != null) parameters.Set("formula2", formula2);
        parameters.Set("backgroundColor", backgroundColor);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit conditional formatting operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="conditionalFormattingIndex">The index of conditional formatting to edit.</param>
    /// <param name="conditionIndex">The condition index within the formatting rule.</param>
    /// <param name="condition">The condition type.</param>
    /// <param name="value">The condition value or first formula.</param>
    /// <param name="formula2">The second value for between condition.</param>
    /// <param name="backgroundColor">The background color for matching cells.</param>
    /// <returns>OperationParameters configured for editing conditional formatting.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters,
        int conditionalFormattingIndex, int? conditionIndex, string? condition, string? value, string? formula2,
        string backgroundColor)
    {
        parameters.Set("conditionalFormattingIndex", conditionalFormattingIndex);
        if (conditionIndex.HasValue) parameters.Set("conditionIndex", conditionIndex.Value);
        if (condition != null) parameters.Set("condition", condition);
        if (value != null) parameters.Set("value", value);
        if (formula2 != null) parameters.Set("formula2", formula2);
        parameters.Set("backgroundColor", backgroundColor);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete conditional formatting operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="conditionalFormattingIndex">The index of conditional formatting to delete.</param>
    /// <returns>OperationParameters configured for deleting conditional formatting.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters,
        int conditionalFormattingIndex)
    {
        parameters.Set("conditionalFormattingIndex", conditionalFormattingIndex);
        return parameters;
    }
}
