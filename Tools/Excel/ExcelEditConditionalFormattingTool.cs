using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditConditionalFormattingTool : IAsposeTool
{
    public string Description => "Edit an existing conditional formatting rule in an Excel worksheet";

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
            conditionalFormattingIndex = new
            {
                type = "number",
                description = "Conditional formatting index to edit (0-based)"
            },
            conditionIndex = new
            {
                type = "number",
                description = "Condition index within the formatting rule (0-based, optional)"
            },
            condition = new
            {
                type = "string",
                description = "New condition type (GreaterThan, LessThan, Between, Equal, optional)"
            },
            value = new
            {
                type = "string",
                description = "New condition value (optional)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "New background color (optional)"
            }
        },
        required = new[] { "path", "conditionalFormattingIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var conditionalFormattingIndex = arguments?["conditionalFormattingIndex"]?.GetValue<int>() ?? throw new ArgumentException("conditionalFormattingIndex is required");
        var conditionIndex = arguments?["conditionIndex"]?.GetValue<int?>();
        var conditionStr = arguments?["condition"]?.GetValue<string>();
        var value = arguments?["value"]?.GetValue<string>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"條件格式索引 {conditionalFormattingIndex} 超出範圍 (工作表共有 {conditionalFormattings.Count} 個條件格式規則)");
        }

        var cf = conditionalFormattings[conditionalFormattingIndex];
        var changes = new List<string>();

        // Edit specific condition or first condition if not specified
        int condIndex = conditionIndex ?? 0;
        if (condIndex >= 0 && condIndex < cf.Count)
        {
            var condition = cf[condIndex];

            // Update condition type and operator
            if (!string.IsNullOrEmpty(conditionStr))
            {
                var operatorType = conditionStr.ToLower() switch
                {
                    "greaterthan" => OperatorType.GreaterThan,
                    "lessthan" => OperatorType.LessThan,
                    "between" => OperatorType.Between,
                    "equal" => OperatorType.Equal,
                    _ => condition.Operator
                };
                condition.Operator = operatorType;
                changes.Add($"條件運算符: {conditionStr}");
            }

            // Update value
            if (!string.IsNullOrEmpty(value))
            {
                condition.Formula1 = value;
                changes.Add($"條件值: {value}");
            }

            // Update background color
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                condition.Style.ForegroundColor = System.Drawing.Color.FromName(backgroundColor);
                condition.Style.Pattern = BackgroundType.Solid;
                changes.Add($"背景顏色: {backgroundColor}");
            }
        }

        workbook.Save(outputPath);

        var result = $"成功編輯條件格式規則 #{conditionalFormattingIndex}";
        if (conditionIndex.HasValue)
        {
            result += $", 條件 #{conditionIndex.Value}";
        }
        result += "\n";
        
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

