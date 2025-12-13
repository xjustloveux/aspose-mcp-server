using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel conditional formatting (add, edit, delete, get)
/// Merges: ExcelAddConditionalFormattingTool, ExcelEditConditionalFormattingTool, 
/// ExcelDeleteConditionalFormattingTool, ExcelGetConditionalFormattingTool
/// </summary>
public class ExcelConditionalFormattingTool : IAsposeTool
{
    public string Description => @"Manage Excel conditional formatting. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add conditional formatting: excel_conditional_formatting(operation='add', path='book.xlsx', range='A1:A10', type='CellValue', operator='Between', formula1='10', formula2='100')
- Edit conditional formatting: excel_conditional_formatting(operation='edit', path='book.xlsx', conditionalFormattingIndex=0, type='CellValue')
- Delete conditional formatting: excel_conditional_formatting(operation='delete', path='book.xlsx', conditionalFormattingIndex=0)
- Get conditional formatting: excel_conditional_formatting(operation='get', path='book.xlsx', range='A1:A10')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add conditional formatting (required params: path, range, type)
- 'edit': Edit conditional formatting (required params: path, conditionalFormattingIndex)
- 'delete': Delete conditional formatting (required params: path, conditionalFormattingIndex)
- 'get': Get conditional formatting (required params: path, range)",
                @enum = new[] { "add", "edit", "delete", "get" }
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
                description = "Cell range (e.g., 'A1:A10', required for add)"
            },
            conditionalFormattingIndex = new
            {
                type = "number",
                description = "Conditional formatting index (0-based, required for edit/delete)"
            },
            conditionIndex = new
            {
                type = "number",
                description = "Condition index within the formatting rule (0-based, optional for edit)"
            },
            condition = new
            {
                type = "string",
                description = "Condition type (GreaterThan, LessThan, Between, Equal, required for add)"
            },
            value = new
            {
                type = "string",
                description = "Condition value (required for add)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color for matching cells (optional)"
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
            "add" => await AddConditionalFormattingAsync(arguments, path, sheetIndex),
            "edit" => await EditConditionalFormattingAsync(arguments, path, sheetIndex),
            "delete" => await DeleteConditionalFormattingAsync(arguments, path, sheetIndex),
            "get" => await GetConditionalFormattingAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for add operation");
        var conditionStr = arguments?["condition"]?.GetValue<string>() ?? throw new ArgumentException("condition is required for add operation");
        var value = arguments?["value"]?.GetValue<string>() ?? throw new ArgumentException("value is required for add operation");
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>() ?? "Yellow";

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        int formatIndex = worksheet.ConditionalFormattings.Add();
        var fcs = worksheet.ConditionalFormattings[formatIndex];
        
        var cellRange = worksheet.Cells.CreateRange(range);
        var areaIndex = fcs.AddArea(new CellArea 
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
        
        // Handle Between condition - need Formula2
        if (operatorType == OperatorType.Between && value.Contains(","))
        {
            var parts = value.Split(',');
            if (parts.Length >= 2)
            {
                fc.Formula1 = parts[0].Trim();
                fc.Formula2 = parts[1].Trim();
            }
            else
            {
                fc.Formula1 = value;
            }
        }
        else
        {
            fc.Formula1 = value;
        }
        
        // Handle both color names (e.g., "Red", "Yellow") and hex values (e.g., "#FF0000")
        // Use BackgroundColor for conditional formatting background
        try
        {
            if (backgroundColor.StartsWith("#"))
            {
                fc.Style.BackgroundColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor);
            }
            else
            {
                fc.Style.BackgroundColor = System.Drawing.Color.FromName(backgroundColor);
            }
        }
        catch
        {
            // Fallback to default color if parsing fails
            fc.Style.BackgroundColor = System.Drawing.Color.Yellow;
        }
        fc.Style.Pattern = BackgroundType.Solid;

        workbook.CalculateFormula();
        workbook.Save(path);

        return await Task.FromResult($"條件格式已添加到範圍 {range} ({conditionStr}): {path}");
    }

    private async Task<string> EditConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var conditionalFormattingIndex = arguments?["conditionalFormattingIndex"]?.GetValue<int>() ?? throw new ArgumentException("conditionalFormattingIndex is required for edit operation");
        var conditionIndex = arguments?["conditionIndex"]?.GetValue<int?>();
        var conditionStr = arguments?["condition"]?.GetValue<string>();
        var value = arguments?["value"]?.GetValue<string>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"條件格式索引 {conditionalFormattingIndex} 超出範圍 (工作表共有 {conditionalFormattings.Count} 個條件格式)");
        }

        var fcs = conditionalFormattings[conditionalFormattingIndex];
        var changes = new List<string>();

        if (conditionIndex.HasValue)
        {
            if (conditionIndex.Value < 0 || conditionIndex.Value >= fcs.Count)
            {
                throw new ArgumentException($"條件索引 {conditionIndex.Value} 超出範圍");
            }

            var condition = fcs[conditionIndex.Value];

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
                changes.Add($"運算符: {conditionStr}");
            }

            if (!string.IsNullOrEmpty(value))
            {
                if (condition.Operator == OperatorType.Between && value.Contains(","))
                {
                    var parts = value.Split(',');
                    if (parts.Length >= 2)
                    {
                        condition.Formula1 = parts[0].Trim();
                        condition.Formula2 = parts[1].Trim();
                    }
                }
                else
                {
                    condition.Formula1 = value;
                }
                changes.Add($"值: {value}");
            }

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                var style = condition.Style;
                // Handle both color names (e.g., "Red", "Yellow") and hex values (e.g., "#FF0000")
                try
                {
                    if (backgroundColor.StartsWith("#"))
                    {
                        style.BackgroundColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor);
                    }
                    else
                    {
                        style.BackgroundColor = System.Drawing.Color.FromName(backgroundColor);
                    }
                }
                catch
                {
                    // Fallback to default color if parsing fails
                    style.BackgroundColor = System.Drawing.Color.Yellow;
                }
                style.Pattern = BackgroundType.Solid;
                changes.Add($"背景色: {backgroundColor}");
            }
        }
        
        // Force recalculation to ensure conditional formatting is applied
        workbook.CalculateFormula();

        workbook.Save(outputPath);

        var result = $"成功編輯條件格式 #{conditionalFormattingIndex}\n";
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

    private async Task<string> DeleteConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var conditionalFormattingIndex = arguments?["conditionalFormattingIndex"]?.GetValue<int>() ?? throw new ArgumentException("conditionalFormattingIndex is required for delete operation");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"條件格式索引 {conditionalFormattingIndex} 超出範圍 (工作表共有 {conditionalFormattings.Count} 個條件格式)");
        }

        conditionalFormattings.RemoveAt(conditionalFormattingIndex);
        workbook.Save(path);
        
        var remainingCount = conditionalFormattings.Count;
        
        return await Task.FromResult($"成功刪除條件格式 #{conditionalFormattingIndex}\n工作表剩餘條件格式數: {remainingCount}\n輸出: {path}");
    }

    private async Task<string> GetConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的條件格式資訊 ===\n");
        result.AppendLine($"總條件格式數: {conditionalFormattings.Count}\n");

        if (conditionalFormattings.Count == 0)
        {
            result.AppendLine("未找到條件格式");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < conditionalFormattings.Count; i++)
        {
            var cf = conditionalFormattings[i];
            result.AppendLine($"【條件格式 {i}】");
            result.AppendLine($"條件格式集合索引: {i}");
            result.AppendLine("狀態: 條件格式已應用");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

}

