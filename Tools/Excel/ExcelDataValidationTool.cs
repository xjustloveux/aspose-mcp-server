using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel data validation (add, edit, delete, get, set messages)
///     Merges: ExcelAddDataValidationTool, ExcelEditDataValidationTool, ExcelDeleteDataValidationTool,
///     ExcelGetDataValidationTool, ExcelSetDataValidationInputMessageTool, ExcelSetDataValidationErrorMessageTool
/// </summary>
public class ExcelDataValidationTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage Excel data validation. Supports 5 operations: add, edit, delete, get, set_messages.

Usage examples:
- Add validation: excel_data_validation(operation='add', path='book.xlsx', range='A1:A10', validationType='List', formula1='1,2,3')
- Edit validation: excel_data_validation(operation='edit', path='book.xlsx', validationIndex=0, validationType='WholeNumber', formula1='0', formula2='100')
- Delete validation: excel_data_validation(operation='delete', path='book.xlsx', validationIndex=0)
- Get validation: excel_data_validation(operation='get', path='book.xlsx', validationIndex=0)
- Set messages: excel_data_validation(operation='set_messages', path='book.xlsx', validationIndex=0, inputMessage='Enter value', errorMessage='Invalid value')";

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
- 'get': Get data validation info (required params: path, validationIndex)
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
                    "Data validation index (0-based, required for edit, delete, get, and set_messages operations)"
            },
            validationType = new
            {
                type = "string",
                description =
                    "Validation type: 'WholeNumber', 'Decimal', 'List', 'Date', 'Time', 'TextLength', 'Custom'",
                @enum = new[] { "WholeNumber", "Decimal", "List", "Date", "Time", "TextLength", "Custom" }
            },
            formula1 = new
            {
                type = "string",
                description = "First formula/value (e.g., '1,2,3' for List, '0' for minimum, required for add)"
            },
            formula2 = new
            {
                type = "string",
                description = "Second formula/value (optional, for range validations like 'between')"
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
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
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
    ///     Adds data validation to a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">
    ///     JSON arguments containing range, validationType, formula1, optional formula2, showError,
    ///     errorTitle, errorMessage
    /// </param>
    /// <returns>Success message</returns>
    private Task<string> AddDataValidationAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var validationType = ArgumentHelper.GetString(arguments, "validationType");
            var formula1 = ArgumentHelper.GetString(arguments, "formula1");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var errorMessage = ArgumentHelper.GetStringNullable(arguments, "errorMessage");
            var inputMessage = ArgumentHelper.GetStringNullable(arguments, "inputMessage");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
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

            var vType = validationType switch
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

            validation.Type = vType;
            validation.Formula1 = formula1;

            if (!string.IsNullOrEmpty(formula2))
            {
                validation.Formula2 = formula2;
                validation.Operator = OperatorType.Between;
            }
            else
            {
                validation.Operator = OperatorType.Equal;
            }

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

            validation.InCellDropDown = true;

            workbook.Save(outputPath);

            return $"Data validation added to range {range} ({validationType}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits existing data validation
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">
    ///     JSON arguments containing range, optional validationType, formula1, formula2, showError,
    ///     errorTitle, errorMessage
    /// </param>
    /// <returns>Success message</returns>
    private Task<string> EditDataValidationAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var validationIndex = ArgumentHelper.GetInt(arguments, "validationIndex");
            var validationTypeStr = ArgumentHelper.GetStringNullable(arguments, "validationType");
            var formula1 = ArgumentHelper.GetStringNullable(arguments, "formula1");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var errorMessage = ArgumentHelper.GetStringNullable(arguments, "errorMessage");
            var inputMessage = ArgumentHelper.GetStringNullable(arguments, "inputMessage");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            if (validationIndex < 0 || validationIndex >= validations.Count)
                throw new ArgumentException(
                    $"Data validation index {validationIndex} is out of range (worksheet has {validations.Count} data validation rules)");

            var validation = validations[validationIndex];
            var changes = new List<string>();

            if (!string.IsNullOrEmpty(validationTypeStr))
            {
                var vType = validationTypeStr switch
                {
                    "WholeNumber" => ValidationType.WholeNumber,
                    "Decimal" => ValidationType.Decimal,
                    "List" => ValidationType.List,
                    "Date" => ValidationType.Date,
                    "Time" => ValidationType.Time,
                    "TextLength" => ValidationType.TextLength,
                    "Custom" => ValidationType.Custom,
                    _ => validation.Type
                };
                validation.Type = vType;
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
                if (!string.IsNullOrEmpty(formula2)) validation.Operator = OperatorType.Between;
                changes.Add($"Formula2={formula2}");
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
    ///     Deletes data validation from a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteDataValidationAsync(string path, string outputPath, int sheetIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var validationIndex = ArgumentHelper.GetInt(arguments, "validationIndex");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            PowerPointHelper.ValidateCollectionIndex(validationIndex, validations, "data validation");

            validations.RemoveAt(validationIndex);
            workbook.Save(outputPath);

            var remainingCount = validations.Count;

            return $"Deleted data validation #{validationIndex} (remaining: {remainingCount}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets data validation information for a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with data validation details</returns>
    private Task<string> GetDataValidationAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var validations = worksheet.Validations;

            if (validations.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    worksheetName = worksheet.Name,
                    items = Array.Empty<object>(),
                    message = "No data validations found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

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
    ///     Sets input/error messages for data validation
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range, optional inputTitle, inputMessage, errorTitle, errorMessage</param>
    /// <returns>Success message</returns>
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

            PowerPointHelper.ValidateCollectionIndex(validationIndex, validations, "data validation");

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
}