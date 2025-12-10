using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddDataValidationTool : IAsposeTool
{
    public string Description => "Add data validation to cells in an Excel worksheet";

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
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to apply validation (e.g., 'A1:A10')"
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
                description = "First formula/value (e.g., '1,2,3' for List, '0' for minimum, etc.)"
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
        required = new[] { "path", "range", "validationType", "formula1" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var validationType = arguments?["validationType"]?.GetValue<string>() ?? throw new ArgumentException("validationType is required");
        var formula1 = arguments?["formula1"]?.GetValue<string>() ?? throw new ArgumentException("formula1 is required");
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
}

