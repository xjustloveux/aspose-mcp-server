using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditDataValidationTool : IAsposeTool
{
    public string Description => "Edit an existing data validation rule in an Excel worksheet";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            validationIndex = new
            {
                type = "number",
                description = "Data validation index to edit (0-based)"
            },
            validationType = new
            {
                type = "string",
                description = "New validation type (WholeNumber, Decimal, List, Date, Time, TextLength, Custom, optional)"
            },
            formula1 = new
            {
                type = "string",
                description = "New first formula/value (optional)"
            },
            formula2 = new
            {
                type = "string",
                description = "New second formula/value (optional)"
            },
            errorMessage = new
            {
                type = "string",
                description = "New error message (optional)"
            },
            inputMessage = new
            {
                type = "string",
                description = "New input message (optional)"
            }
        },
        required = new[] { "path", "validationIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var validationIndex = arguments?["validationIndex"]?.GetValue<int>() ?? throw new ArgumentException("validationIndex is required");
        var validationTypeStr = arguments?["validationType"]?.GetValue<string>();
        var formula1 = arguments?["formula1"]?.GetValue<string>();
        var formula2 = arguments?["formula2"]?.GetValue<string>();
        var errorMessage = arguments?["errorMessage"]?.GetValue<string>();
        var inputMessage = arguments?["inputMessage"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var validations = worksheet.Validations;
        
        if (validationIndex < 0 || validationIndex >= validations.Count)
        {
            throw new ArgumentException($"數據驗證索引 {validationIndex} 超出範圍 (工作表共有 {validations.Count} 個數據驗證規則)");
        }

        var validation = validations[validationIndex];
        var changes = new List<string>();

        // Update validation type
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

        // Update formula1
        if (!string.IsNullOrEmpty(formula1))
        {
            validation.Formula1 = formula1;
            changes.Add($"公式1: {formula1}");
        }

        // Update formula2
        if (formula2 != null) // Allow empty string to clear formula2
        {
            validation.Formula2 = formula2;
            if (!string.IsNullOrEmpty(formula2))
            {
                validation.Operator = OperatorType.Between;
            }
            changes.Add($"公式2: {formula2 ?? "(已清除)"}");
        }

        // Update error message
        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"錯誤訊息: {errorMessage ?? "(已清除)"}");
        }

        // Update input message
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
}

