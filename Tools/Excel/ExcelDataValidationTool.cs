using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
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
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelDataValidationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelDataValidationTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

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
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddDataValidation(ctx, outputPath, sheetIndex, range, validationType, formula1, formula2,
                operatorType, inCellDropDown, errorMessage, inputMessage),
            "edit" => EditDataValidation(ctx, outputPath, sheetIndex, validationIndex, validationType, formula1,
                formula2, operatorType, inCellDropDown, errorMessage, inputMessage),
            "delete" => DeleteDataValidation(ctx, outputPath, sheetIndex, validationIndex),
            "get" => GetDataValidation(ctx, sheetIndex),
            "set_messages" => SetMessages(ctx, outputPath, sheetIndex, validationIndex, errorMessage, inputMessage),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds data validation to a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="range">The cell range to apply validation (e.g., 'A1:A10').</param>
    /// <param name="validationType">The type of validation (e.g., 'WholeNumber', 'List').</param>
    /// <param name="formula1">The first formula/value for the validation.</param>
    /// <param name="formula2">The second formula/value (for 'Between' operator).</param>
    /// <param name="operatorType">The operator type (e.g., 'Between', 'Equal').</param>
    /// <param name="inCellDropDown">Whether to show dropdown list in cell (for List type).</param>
    /// <param name="errorMessage">The error message to show when validation fails.</param>
    /// <param name="inputMessage">The input message to show when cell is selected.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private static string AddDataValidation(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, string? validationType, string? formula1, string? formula2, string? operatorType,
        bool inCellDropDown, string? errorMessage, string? inputMessage)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for add operation");
        if (string.IsNullOrEmpty(validationType))
            throw new ArgumentException("validationType is required for add operation");
        if (string.IsNullOrEmpty(formula1))
            throw new ArgumentException("formula1 is required for add operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        var area = new CellArea
        {
            StartRow = cellRange.FirstRow,
            StartColumn = cellRange.FirstColumn,
            EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
            EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
        };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];

        var vType = ParseValidationType(validationType);
        validation.Type = vType;
        validation.Formula1 = formula1;

        if (!string.IsNullOrEmpty(formula2))
            validation.Formula2 = formula2;

        validation.Operator = ParseOperatorType(operatorType, formula2);

        if (!string.IsNullOrEmpty(errorMessage))
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = true;
        }

        if (!string.IsNullOrEmpty(inputMessage))
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = true;
        }

        if (vType == ValidationType.List)
            validation.InCellDropDown = inCellDropDown;

        ctx.Save(outputPath);

        return
            $"Data validation added to range {range} (type: {validationType}, index: {validationIndex}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits existing data validation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="validationIndex">The data validation index (0-based).</param>
    /// <param name="validationType">The type of validation to set.</param>
    /// <param name="formula1">The first formula/value to set.</param>
    /// <param name="formula2">The second formula/value to set.</param>
    /// <param name="operatorType">The operator type to set.</param>
    /// <param name="inCellDropDown">Whether to show dropdown list in cell.</param>
    /// <param name="errorMessage">The error message to set.</param>
    /// <param name="inputMessage">The input message to set.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the validation index is out of range.</exception>
    private static string EditDataValidation(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int validationIndex, string? validationType, string? formula1, string? formula2, string? operatorType,
        bool? inCellDropDown, string? errorMessage, string? inputMessage)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

        var validation = validations[validationIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(validationType))
        {
            validation.Type = ParseValidationType(validationType);
            changes.Add($"Type={validationType}");
        }

        if (!string.IsNullOrEmpty(formula1))
        {
            validation.Formula1 = formula1;
            changes.Add($"Formula1={formula1}");
        }

        if (formula2 != null)
        {
            validation.Formula2 = formula2;
            changes.Add($"Formula2={formula2}");
        }

        if (!string.IsNullOrEmpty(operatorType))
        {
            validation.Operator = ParseOperatorType(operatorType, formula2);
            changes.Add($"Operator={operatorType}");
        }

        if (inCellDropDown.HasValue)
        {
            validation.InCellDropDown = inCellDropDown.Value;
            changes.Add($"InCellDropDown={inCellDropDown.Value}");
        }

        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"ErrorMessage={errorMessage}");
        }

        if (inputMessage != null)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(inputMessage);
            changes.Add($"InputMessage={inputMessage}");
        }

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
        return $"Edited data validation #{validationIndex} ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes data validation by index.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="validationIndex">The data validation index (0-based) to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the validation index is out of range.</exception>
    private static string DeleteDataValidation(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int validationIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

        validations.RemoveAt(validationIndex);

        ctx.Save(outputPath);

        return
            $"Deleted data validation #{validationIndex} (remaining: {validations.Count}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all data validation information for the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <returns>A JSON string containing all data validation information.</returns>
    private static string GetDataValidation(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        List<object> validationList = [];
        for (var i = 0; i < validations.Count; i++)
        {
            var validation = validations[i];
            validationList.Add(new
            {
                index = i,
                type = validation.Type.ToString(),
                operatorType = validation.Operator.ToString(),
                formula1 = validation.Formula1,
                formula2 = validation.Formula2,
                errorMessage = validation.ErrorMessage,
                inputMessage = validation.InputMessage,
                showError = validation.ShowError,
                showInput = validation.ShowInput,
                inCellDropDown = validation.InCellDropDown
            });
        }

        var result = new
        {
            count = validations.Count,
            worksheetName = worksheet.Name,
            items = validationList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Sets input/error messages for data validation.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="validationIndex">The data validation index (0-based).</param>
    /// <param name="errorMessage">The error message to set.</param>
    /// <param name="inputMessage">The input message to set.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the validation index is out of range.</exception>
    private static string SetMessages(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int validationIndex, string? errorMessage, string? inputMessage)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

        var validation = validations[validationIndex];
        List<string> changes = [];

        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"ErrorMessage={errorMessage}");
        }

        if (inputMessage != null)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(inputMessage);
            changes.Add($"InputMessage={inputMessage}");
        }

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
        return
            $"Updated data validation #{validationIndex} messages ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Parses validation type string to ValidationType enum.
    /// </summary>
    /// <param name="validationType">The validation type string to parse.</param>
    /// <returns>The corresponding ValidationType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the validation type is not supported.</exception>
    private static ValidationType ParseValidationType(string validationType)
    {
        return validationType switch
        {
            "WholeNumber" => ValidationType.WholeNumber,
            "Decimal" => ValidationType.Decimal,
            "List" => ValidationType.List,
            "Date" => ValidationType.Date,
            "Time" => ValidationType.Time,
            "TextLength" => ValidationType.TextLength,
            "Custom" => ValidationType.Custom,
            _ => throw new ArgumentException($"Unsupported validation type: {validationType}")
        };
    }

    /// <summary>
    ///     Parses operator type string to OperatorType enum.
    /// </summary>
    /// <param name="operatorType">The operator type string to parse.</param>
    /// <param name="formula2">The second formula value used to infer operator type if not specified.</param>
    /// <returns>The corresponding OperatorType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the operator type is not supported.</exception>
    private static OperatorType ParseOperatorType(string? operatorType, string? formula2)
    {
        if (!string.IsNullOrEmpty(operatorType))
            return operatorType switch
            {
                "Between" => OperatorType.Between,
                "Equal" => OperatorType.Equal,
                "NotEqual" => OperatorType.NotEqual,
                "GreaterThan" => OperatorType.GreaterThan,
                "LessThan" => OperatorType.LessThan,
                "GreaterOrEqual" => OperatorType.GreaterOrEqual,
                "LessOrEqual" => OperatorType.LessOrEqual,
                _ => throw new ArgumentException($"Unsupported operator type: {operatorType}")
            };

        return !string.IsNullOrEmpty(formula2) ? OperatorType.Between : OperatorType.Equal;
    }

    /// <summary>
    ///     Validates collection index and throws exception if invalid.
    /// </summary>
    /// <param name="index">The index to validate.</param>
    /// <param name="count">The total count of items in the collection.</param>
    /// <param name="itemName">The name of the item type for error messages.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of range.</exception>
    private static void ValidateCollectionIndex(int index, int count, string itemName)
    {
        if (index < 0 || index >= count)
            throw new ArgumentException(
                $"{itemName} index {index} is out of range (collection has {count} {itemName}s)");
    }
}