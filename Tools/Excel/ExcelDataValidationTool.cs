using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel data validation (add, edit, delete, get, set messages).
///     Merges: ExcelAddDataValidationTool, ExcelEditDataValidationTool, ExcelDeleteDataValidationTool,
///     ExcelGetDataValidationTool, ExcelSetDataValidationInputMessageTool, ExcelSetDataValidationErrorMessageTool.
/// </summary>
public class ExcelDataValidationTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel data validation. Supports 5 operations: add, edit, delete, get, set_messages.

Usage examples:
- Add list validation: excel_data_validation(operation='add', path='book.xlsx', range='A1:A10', validationType='List', formula1='1,2,3')
- Add number range: excel_data_validation(operation='add', path='book.xlsx', range='B1:B10', validationType='WholeNumber', operatorType='Between', formula1='0', formula2='100')
- Add greater than: excel_data_validation(operation='add', path='book.xlsx', range='C1:C10', validationType='WholeNumber', operatorType='GreaterThan', formula1='0')
- Edit validation: excel_data_validation(operation='edit', path='book.xlsx', validationIndex=0, validationType='WholeNumber', formula1='0', formula2='100')
- Delete validation: excel_data_validation(operation='delete', path='book.xlsx', validationIndex=0)
- Get validation: excel_data_validation(operation='get', path='book.xlsx')
- Set messages: excel_data_validation(operation='set_messages', path='book.xlsx', validationIndex=0, inputMessage='Enter value', errorMessage='Invalid value')";

    /// <summary>
    ///     Gets the JSON schema for the tool's input parameters.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add data validation (required params: path, range, validationType, formula1)
- 'edit': Edit data validation (required params: path, validationIndex)
- 'delete': Delete data validation (required params: path, validationIndex)
- 'get': Get data validation info (required params: path)
- 'set_messages': Set input/error messages (required params: path, validationIndex)",
                @enum = new[] { "add", "edit", "delete", "get", "set_messages" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for add/edit/delete/set_messages operations, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to apply validation (e.g., 'A1:A10', required for add operation)"
            },
            validationIndex = new
            {
                type = "number",
                description =
                    "Data validation index (0-based, required for edit, delete, and set_messages operations)"
            },
            validationType = new
            {
                type = "string",
                description =
                    "Validation type: 'WholeNumber', 'Decimal', 'List', 'Date', 'Time', 'TextLength', 'Custom'",
                @enum = new[] { "WholeNumber", "Decimal", "List", "Date", "Time", "TextLength", "Custom" }
            },
            operatorType = new
            {
                type = "string",
                description =
                    "Operator type for validation comparison. Default: 'Between' if formula2 is provided, 'Equal' otherwise. For List type, this is ignored.",
                @enum = new[]
                    { "Between", "Equal", "NotEqual", "GreaterThan", "LessThan", "GreaterOrEqual", "LessOrEqual" }
            },
            formula1 = new
            {
                type = "string",
                description = "First formula/value (e.g., '1,2,3' for List, '0' for minimum, required for add)"
            },
            formula2 = new
            {
                type = "string",
                description =
                    "Second formula/value (required for 'Between' operator, optional for other operators)"
            },
            inCellDropDown = new
            {
                type = "boolean",
                description =
                    "Show dropdown list in cell (only applicable for List type, optional, default: true)"
            },
            errorMessage = new
            {
                type = "string",
                description = "Error message to show when validation fails (optional)"
            },
            inputMessage = new
            {
                type = "string",
                description = "Input message to show when cell is selected (optional)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddDataValidationAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditDataValidationAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteDataValidationAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetDataValidationAsync(path, sheetIndex),
            "set_messages" => await SetMessagesAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds data validation to a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">
    ///     JSON arguments containing range, validationType, formula1, optional formula2, operatorType,
    ///     inCellDropDown, errorMessage, inputMessage.
    /// </param>
    /// <returns>Success message with validation index.</returns>
    private Task<string> AddDataValidationAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var validationType = ArgumentHelper.GetString(arguments, "validationType");
            var formula1 = ArgumentHelper.GetString(arguments, "formula1");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var operatorTypeStr = ArgumentHelper.GetStringNullable(arguments, "operatorType");
            var inCellDropDown = ArgumentHelper.GetBool(arguments, "inCellDropDown", true);
            var errorMessage = ArgumentHelper.GetStringNullable(arguments, "errorMessage");
            var inputMessage = ArgumentHelper.GetStringNullable(arguments, "inputMessage");

            using var workbook = new Workbook(path);
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

            validation.Operator = ParseOperatorType(operatorTypeStr, formula2);

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

            workbook.Save(outputPath);

            return
                $"Data validation added to range {range} (type: {validationType}, index: {validationIndex}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits existing data validation.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">
    ///     JSON arguments containing validationIndex, optional validationType, formula1, formula2,
    ///     operatorType, inCellDropDown, errorMessage, inputMessage.
    /// </param>
    /// <returns>Success message with changes made.</returns>
    private Task<string> EditDataValidationAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var validationIndex = ArgumentHelper.GetInt(arguments, "validationIndex");
            var validationTypeStr = ArgumentHelper.GetStringNullable(arguments, "validationType");
            var formula1 = ArgumentHelper.GetStringNullable(arguments, "formula1");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var operatorTypeStr = ArgumentHelper.GetStringNullable(arguments, "operatorType");
            var inCellDropDown = ArgumentHelper.GetBoolNullable(arguments, "inCellDropDown");
            var errorMessage = ArgumentHelper.GetStringNullable(arguments, "errorMessage");
            var inputMessage = ArgumentHelper.GetStringNullable(arguments, "inputMessage");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

            var validation = validations[validationIndex];
            var changes = new List<string>();

            if (!string.IsNullOrEmpty(validationTypeStr))
            {
                validation.Type = ParseValidationType(validationTypeStr);
                changes.Add($"Type={validationTypeStr}");
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

            if (!string.IsNullOrEmpty(operatorTypeStr))
            {
                validation.Operator = ParseOperatorType(operatorTypeStr, formula2);
                changes.Add($"Operator={operatorTypeStr}");
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

            workbook.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
            return $"Edited data validation #{validationIndex} ({changesStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes data validation by index.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing validationIndex.</param>
    /// <returns>Success message with remaining count.</returns>
    private Task<string> DeleteDataValidationAsync(string path, string outputPath, int sheetIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var validationIndex = ArgumentHelper.GetInt(arguments, "validationIndex");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

            validations.RemoveAt(validationIndex);
            workbook.Save(outputPath);

            return $"Deleted data validation #{validationIndex} (remaining: {validations.Count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all data validation information for the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with data validation details.</returns>
    private Task<string> GetDataValidationAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            var validationList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Sets input/error messages for data validation.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing validationIndex, optional inputMessage, errorMessage.</param>
    /// <returns>Success message with changes made.</returns>
    private Task<string> SetMessagesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var validationIndex = ArgumentHelper.GetInt(arguments, "validationIndex");
            var errorMessage = ArgumentHelper.GetStringNullable(arguments, "errorMessage");
            var inputMessage = ArgumentHelper.GetStringNullable(arguments, "inputMessage");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

            var validation = validations[validationIndex];
            var changes = new List<string>();

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

            workbook.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
            return $"Updated data validation #{validationIndex} messages ({changesStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Parses validation type string to ValidationType enum.
    /// </summary>
    /// <param name="validationType">Validation type string.</param>
    /// <returns>ValidationType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown if validation type is not supported.</exception>
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
    /// <param name="operatorType">Operator type string (can be null).</param>
    /// <param name="formula2">Second formula value (used for default determination).</param>
    /// <returns>OperatorType enum value.</returns>
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
    /// <param name="index">Index to validate.</param>
    /// <param name="count">Collection count.</param>
    /// <param name="itemName">Name of the item type for error message.</param>
    /// <exception cref="ArgumentException">Thrown if index is out of range.</exception>
    private static void ValidateCollectionIndex(int index, int count, string itemName)
    {
        if (index < 0 || index >= count)
            throw new ArgumentException(
                $"{itemName} index {index} is out of range (collection has {count} {itemName}s)");
    }
}