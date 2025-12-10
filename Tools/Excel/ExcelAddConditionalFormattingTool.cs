using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddConditionalFormattingTool : IAsposeTool
{
    public string Description => "Add conditional formatting to Excel cells";

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
                description = "Cell range (e.g., 'A1:A10')"
            },
            condition = new
            {
                type = "string",
                description = "Condition type (GreaterThan, LessThan, Between, etc.)"
            },
            value = new
            {
                type = "string",
                description = "Condition value"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color for matching cells"
            }
        },
        required = new[] { "path", "range", "condition", "value" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var conditionStr = arguments?["condition"]?.GetValue<string>() ?? throw new ArgumentException("condition is required");
        var value = arguments?["value"]?.GetValue<string>() ?? throw new ArgumentException("value is required");
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>() ?? "Yellow";

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        int formatIndex = worksheet.ConditionalFormattings.Add();
        var fcs = worksheet.ConditionalFormattings[formatIndex];
        
        var cellRange = worksheet.Cells.CreateRange(range);
        var ca = fcs.AddArea(new CellArea 
        { 
            StartRow = cellRange.FirstRow,
            EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
            StartColumn = cellRange.FirstColumn,
            EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
        });

        var conditionType = conditionStr.ToLower() switch
        {
            "greaterthan" => FormatConditionType.CellValue,
            "lessthan" => FormatConditionType.CellValue,
            "between" => FormatConditionType.CellValue,
            "equal" => FormatConditionType.CellValue,
            _ => FormatConditionType.CellValue
        };

        int conditionIndex = fcs.AddCondition(conditionType);
        var fc = fcs[conditionIndex];

        var operatorType = conditionStr.ToLower() switch
        {
            "greaterthan" => OperatorType.GreaterThan,
            "lessthan" => OperatorType.LessThan,
            "between" => OperatorType.Between,
            "equal" => OperatorType.Equal,
            _ => OperatorType.GreaterThan
        };

        fc.Operator = operatorType;
        fc.Formula1 = value;
        fc.Style.ForegroundColor = System.Drawing.Color.FromName(backgroundColor);
        fc.Style.Pattern = BackgroundType.Solid;

        workbook.Save(path);

        return await Task.FromResult($"Conditional formatting added to range {range}: {path}");
    }
}

