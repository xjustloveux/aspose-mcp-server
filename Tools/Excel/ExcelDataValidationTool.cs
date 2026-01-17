using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel data validation (add, edit, delete, get, set messages)
/// </summary>
[McpServerToolType]
public class ExcelDataValidationTool
{
    /// <summary>
    ///     Handler registry for data validation operations.
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
    ///     Initializes a new instance of the <see cref="ExcelDataValidationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelDataValidationTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.DataValidation");
    }

    /// <summary>
    ///     Executes an Excel data validation operation (add, edit, delete, get, or set_messages).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, or set_messages.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range to apply validation (e.g., 'A1:A10', required for add).</param>
    /// <param name="validationIndex">Data validation index (0-based, required for edit/delete/set_messages).</param>
    /// <param name="validationType">Validation type: WholeNumber, Decimal, List, Date, Time, TextLength, Custom.</param>
    /// <param name="operatorType">Operator type: Between, Equal, NotEqual, GreaterThan, LessThan, GreaterOrEqual, LessOrEqual.</param>
    /// <param name="formula1">First formula/value (e.g., '1,2,3' for List, '0' for minimum, required for add).</param>
    /// <param name="formula2">Second formula/value (required for 'Between' operator).</param>
    /// <param name="inCellDropDown">Show dropdown list in cell (only for List type).</param>
    /// <param name="errorMessage">Error message to show when validation fails.</param>
    /// <param name="inputMessage">Input message to show when cell is selected.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_data_validation")]
    [Description(@"Manage Excel data validation. Supports 5 operations: add, edit, delete, get, set_messages.

Usage examples:
- Add list validation: excel_data_validation(operation='add', path='book.xlsx', range='A1:A10', validationType='List', formula1='1,2,3')
- Add number range: excel_data_validation(operation='add', path='book.xlsx', range='B1:B10', validationType='WholeNumber', operatorType='Between', formula1='0', formula2='100')
- Add greater than: excel_data_validation(operation='add', path='book.xlsx', range='C1:C10', validationType='WholeNumber', operatorType='GreaterThan', formula1='0')
- Edit validation: excel_data_validation(operation='edit', path='book.xlsx', validationIndex=0, validationType='WholeNumber', formula1='0', formula2='100')
- Delete validation: excel_data_validation(operation='delete', path='book.xlsx', validationIndex=0)
- Get validation: excel_data_validation(operation='get', path='book.xlsx')
- Set messages: excel_data_validation(operation='set_messages', path='book.xlsx', validationIndex=0, inputMessage='Enter value', errorMessage='Invalid value')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get, set_messages")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range to apply validation (e.g., 'A1:A10', required for add)")]
        string? range = null,
        [Description("Data validation index (0-based, required for edit/delete/set_messages)")]
        int validationIndex = 0,
        [Description("Validation type: WholeNumber, Decimal, List, Date, Time, TextLength, Custom")]
        string? validationType = null,
        [Description("Operator type: Between, Equal, NotEqual, GreaterThan, LessThan, GreaterOrEqual, LessOrEqual")]
        string? operatorType = null,
        [Description("First formula/value (e.g., '1,2,3' for List, '0' for minimum, required for add)")]
        string? formula1 = null,
        [Description("Second formula/value (required for 'Between' operator)")]
        string? formula2 = null,
        [Description("Show dropdown list in cell (only for List type, default: true)")]
        bool inCellDropDown = true,
        [Description("Error message to show when validation fails (optional)")]
        string? errorMessage = null,
        [Description("Input message to show when cell is selected (optional)")]
        string? inputMessage = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, validationIndex, validationType,
            operatorType, formula1, formula2, inCellDropDown, errorMessage, inputMessage);

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
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? range,
        int validationIndex,
        string? validationType,
        string? operatorType,
        string? formula1,
        string? formula2,
        bool inCellDropDown,
        string? errorMessage,
        string? inputMessage)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, range, validationType, operatorType, formula1, formula2,
                inCellDropDown, errorMessage, inputMessage),
            "edit" => BuildEditParameters(parameters, validationIndex, validationType, operatorType, formula1, formula2,
                inCellDropDown, errorMessage, inputMessage),
            "delete" => BuildDeleteParameters(parameters, validationIndex),
            "get" => parameters,
            "set_messages" => BuildSetMessagesParameters(parameters, validationIndex, errorMessage, inputMessage),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add validation operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range to apply validation.</param>
    /// <param name="validationType">The validation type.</param>
    /// <param name="operatorType">The operator type for validation.</param>
    /// <param name="formula1">The first formula or value.</param>
    /// <param name="formula2">The second formula or value for between operator.</param>
    /// <param name="inCellDropDown">Whether to show dropdown in cell.</param>
    /// <param name="errorMessage">The error message to display.</param>
    /// <param name="inputMessage">The input message to display.</param>
    /// <returns>OperationParameters configured for adding validation.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? range,
        string? validationType, string? operatorType, string? formula1, string? formula2, bool inCellDropDown,
        string? errorMessage, string? inputMessage)
    {
        if (range != null) parameters.Set("range", range);
        if (validationType != null) parameters.Set("validationType", validationType);
        if (formula1 != null) parameters.Set("formula1", formula1);
        if (formula2 != null) parameters.Set("formula2", formula2);
        if (operatorType != null) parameters.Set("operatorType", operatorType);
        parameters.Set("inCellDropDown", inCellDropDown);
        if (errorMessage != null) parameters.Set("errorMessage", errorMessage);
        if (inputMessage != null) parameters.Set("inputMessage", inputMessage);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit validation operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="validationIndex">The index of validation to edit.</param>
    /// <param name="validationType">The validation type.</param>
    /// <param name="operatorType">The operator type for validation.</param>
    /// <param name="formula1">The first formula or value.</param>
    /// <param name="formula2">The second formula or value for between operator.</param>
    /// <param name="inCellDropDown">Whether to show dropdown in cell.</param>
    /// <param name="errorMessage">The error message to display.</param>
    /// <param name="inputMessage">The input message to display.</param>
    /// <returns>OperationParameters configured for editing validation.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int validationIndex,
        string? validationType, string? operatorType, string? formula1, string? formula2, bool inCellDropDown,
        string? errorMessage, string? inputMessage)
    {
        parameters.Set("validationIndex", validationIndex);
        if (validationType != null) parameters.Set("validationType", validationType);
        if (formula1 != null) parameters.Set("formula1", formula1);
        if (formula2 != null) parameters.Set("formula2", formula2);
        if (operatorType != null) parameters.Set("operatorType", operatorType);
        parameters.Set("inCellDropDown", inCellDropDown);
        if (errorMessage != null) parameters.Set("errorMessage", errorMessage);
        if (inputMessage != null) parameters.Set("inputMessage", inputMessage);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete validation operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="validationIndex">The index of validation to delete.</param>
    /// <returns>OperationParameters configured for deleting validation.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int validationIndex)
    {
        parameters.Set("validationIndex", validationIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set messages operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="validationIndex">The index of validation to set messages for.</param>
    /// <param name="errorMessage">The error message to display.</param>
    /// <param name="inputMessage">The input message to display.</param>
    /// <returns>OperationParameters configured for setting validation messages.</returns>
    private static OperationParameters BuildSetMessagesParameters(OperationParameters parameters, int validationIndex,
        string? errorMessage, string? inputMessage)
    {
        parameters.Set("validationIndex", validationIndex);
        if (errorMessage != null) parameters.Set("errorMessage", errorMessage);
        if (inputMessage != null) parameters.Set("inputMessage", inputMessage);
        return parameters;
    }
}
