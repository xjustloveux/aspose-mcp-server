using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel data validation (add, edit, delete, get, set messages)
/// Merges: ExcelAddDataValidationTool, ExcelEditDataValidationTool, ExcelDeleteDataValidationTool, 
/// ExcelGetDataValidationTool, ExcelSetDataValidationInputMessageTool, ExcelSetDataValidationErrorMessageTool
/// </summary>
public class ExcelDataValidationTool : IAsposeTool
{
    public string Description => @"Manage Excel data validation. Supports 5 operations: add, edit, delete, get, set_messages.

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
                description = "Output file path (optional, for edit operation, defaults to input path)"
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
                description = "Data validation index (0-based, required for edit, delete, get, and set_messages operations)"
            },
            validationType = new
            {
                type = "string",
                description = "Validation type: 'WholeNumber', 'Decimal', 'List', 'Date', 'Time', 'TextLength', 'Custom'",
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddDataValidationAsync(arguments, path, sheetIndex),
            "edit" => await EditDataValidationAsync(arguments, path, sheetIndex),
            "delete" => await DeleteDataValidationAsync(arguments, path, sheetIndex),
            "get" => await GetDataValidationAsync(arguments, path, sheetIndex),
            "set_messages" => await SetMessagesAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddDataValidationAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for add operation");
        var validationType = arguments?["validationType"]?.GetValue<string>() ?? throw new ArgumentException("validationType is required for add operation");
        var formula1 = arguments?["formula1"]?.GetValue<string>() ?? throw new ArgumentException("formula1 is required for add operation");
        var formula2 = arguments?["formula2"]?.GetValue<string>();
        var errorMessage = arguments?["errorMessage"]?.GetValue<string>();
        var inputMessage = arguments?["inputMessage"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var area = new CellArea();
        area.StartRow = cellRange.FirstRow;
        area.StartColumn = cellRange.FirstColumn;
        area.EndRow = cellRange.FirstRow + cellRange.RowCount - 1;
        area.EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1;
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

        workbook.Save(path);

        return await Task.FromResult($"範圍 {range} 已添加數據驗證 ({validationType}): {path}");
    }

    private async Task<string> EditDataValidationAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var validationIndex = arguments?["validationIndex"]?.GetValue<int>() ?? throw new ArgumentException("validationIndex is required for edit operation");
        var validationTypeStr = arguments?["validationType"]?.GetValue<string>();
        var formula1 = arguments?["formula1"]?.GetValue<string>();
        var formula2 = arguments?["formula2"]?.GetValue<string>();
        var errorMessage = arguments?["errorMessage"]?.GetValue<string>();
        var inputMessage = arguments?["inputMessage"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;
        
        if (validationIndex < 0 || validationIndex >= validations.Count)
        {
            throw new ArgumentException($"數據驗證索引 {validationIndex} 超出範圍 (工作表共有 {validations.Count} 個數據驗證規則)");
        }

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
            changes.Add($"驗證類型: {validationTypeStr}");
        }

        if (!string.IsNullOrEmpty(formula1))
        {
            validation.Formula1 = formula1;
            changes.Add($"公式1: {formula1}");
        }

        if (formula2 != null)
        {
            validation.Formula2 = formula2;
            if (!string.IsNullOrEmpty(formula2))
            {
                validation.Operator = OperatorType.Between;
            }
            changes.Add($"公式2: {formula2 ?? "(已清除)"}");
        }

        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"錯誤訊息: {errorMessage ?? "(已清除)"}");
        }

        if (inputMessage != null)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(inputMessage);
            changes.Add($"輸入訊息: {inputMessage ?? "(已清除)"}");
        }

        workbook.Save(outputPath);

        var result = $"成功編輯數據驗證規則 #{validationIndex}\n";
        if (changes.Count > 0)
        {
            result += "變更:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "無變更。\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> DeleteDataValidationAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var validationIndex = arguments?["validationIndex"]?.GetValue<int>() ?? throw new ArgumentException("validationIndex is required for delete operation");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;
        
        PowerPointHelper.ValidateCollectionIndex(validationIndex, validations, "數據驗證");

        validations.RemoveAt(validationIndex);
        workbook.Save(path);
        
        var remainingCount = validations.Count;
        
        return await Task.FromResult($"成功刪除數據驗證 #{validationIndex}\n工作表剩餘數據驗證數: {remainingCount}\n輸出: {path}");
    }

    private async Task<string> GetDataValidationAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的數據驗證資訊 ===\n");
        result.AppendLine($"總數據驗證數: {validations.Count}\n");

        if (validations.Count == 0)
        {
            result.AppendLine("未找到數據驗證");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < validations.Count; i++)
        {
            var validation = validations[i];
            result.AppendLine($"【數據驗證 {i}】");
            result.AppendLine($"類型: {validation.Type}");
            result.AppendLine($"運算符: {validation.Operator}");
            result.AppendLine("應用範圍: 已應用（詳細範圍資訊需通過其他方式獲取）");
            
            if (!string.IsNullOrEmpty(validation.Formula1))
            {
                result.AppendLine($"公式1: {validation.Formula1}");
            }
            if (!string.IsNullOrEmpty(validation.Formula2))
            {
                result.AppendLine($"公式2: {validation.Formula2}");
            }
            if (!string.IsNullOrEmpty(validation.ErrorMessage))
            {
                result.AppendLine($"錯誤訊息: {validation.ErrorMessage}");
            }
            if (!string.IsNullOrEmpty(validation.InputMessage))
            {
                result.AppendLine($"輸入訊息: {validation.InputMessage}");
            }
            result.AppendLine($"顯示錯誤: {validation.ShowError}");
            result.AppendLine($"顯示輸入: {validation.ShowInput}");
            result.AppendLine($"下拉列表: {validation.InCellDropDown}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> SetMessagesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var validationIndex = arguments?["validationIndex"]?.GetValue<int>() ?? throw new ArgumentException("validationIndex is required for set_messages operation");
        var errorMessage = arguments?["errorMessage"]?.GetValue<string>();
        var inputMessage = arguments?["inputMessage"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;
        
        PowerPointHelper.ValidateCollectionIndex(validationIndex, validations, "數據驗證");

        var validation = validations[validationIndex];
        var changes = new List<string>();

        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"錯誤訊息: {errorMessage ?? "(已清除)"}");
        }

        if (inputMessage != null)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(inputMessage);
            changes.Add($"輸入訊息: {inputMessage ?? "(已清除)"}");
        }

        workbook.Save(path);

        var result = changes.Count > 0 
            ? $"數據驗證訊息已更新: {string.Join(", ", changes)}\n輸出: {path}"
            : $"無變更\n輸出: {path}";
        
        return await Task.FromResult(result);
    }
}

