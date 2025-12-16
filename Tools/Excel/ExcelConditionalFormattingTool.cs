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

You can add multiple conditional formatting rules to the same range by calling the 'add' operation multiple times. Each rule is independent and will be evaluated separately. To add multiple rules, simply call the 'add' operation multiple times with different conditions for the same range.

Usage examples:
- Add conditional formatting: excel_conditional_formatting(operation='add', path='book.xlsx', range='A1:A10', condition='Between', value='10', formula2='100', backgroundColor='#FF0000')
- Add multiple rules: Call 'add' multiple times with different conditions to create multiple rules for the same range
- Edit conditional formatting: excel_conditional_formatting(operation='edit', path='book.xlsx', conditionalFormattingIndex=0, condition='GreaterThan', value='50')
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
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddConditionalFormattingAsync(arguments, path, sheetIndex),
            "edit" => await EditConditionalFormattingAsync(arguments, path, sheetIndex),
            "delete" => await DeleteConditionalFormattingAsync(arguments, path, sheetIndex),
            "get" => await GetConditionalFormattingAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds conditional formatting to a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, condition, value, optional style properties</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range");
        var conditionStr = ArgumentHelper.GetString(arguments, "condition");
        var value = ArgumentHelper.GetString(arguments, "value");
        var backgroundColor = ArgumentHelper.GetString(arguments, "backgroundColor", "Yellow");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        int formatIndex = worksheet.ConditionalFormattings.Add();
        var fcs = worksheet.ConditionalFormattings[formatIndex];
        
        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
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

        var validConditions = new[] { "greaterthan", "lessthan", "between", "equal" };
        var conditionLower = conditionStr.ToLower();
        var isValidCondition = validConditions.Contains(conditionLower);
        
        var operatorType = conditionLower switch
        {
            "greaterthan" => OperatorType.GreaterThan,
            "lessthan" => OperatorType.LessThan,
            "between" => OperatorType.Between,
            "equal" => OperatorType.Equal,
            _ => OperatorType.GreaterThan
        };

        fc.Operator = operatorType;
        
        // Warning for invalid condition types
        string? warningMessage = null;
        if (!isValidCondition)
        {
            warningMessage = $"\n⚠️ Warning: Condition type '{conditionStr}' may not be supported. Valid types are: GreaterThan, LessThan, Between, Equal. Please verify in Excel.";
        }
        
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
            fc.Style.BackgroundColor = ColorHelper.ParseColor(backgroundColor);
        }
        catch
        {
            // Fallback to default color if parsing fails
            fc.Style.BackgroundColor = System.Drawing.Color.Yellow;
        }
        fc.Style.Pattern = BackgroundType.Solid;

        workbook.CalculateFormula();
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var result = $"Conditional formatting added to range {range} ({conditionStr}): {outputPath}";
        if (!string.IsNullOrEmpty(warningMessage))
        {
            result += warningMessage;
        }
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Edits existing conditional formatting
    /// </summary>
    /// <param name="arguments">JSON arguments containing index, optional condition, value, style properties</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> EditConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var conditionalFormattingIndex = ArgumentHelper.GetInt(arguments, "conditionalFormattingIndex");
        var conditionIndex = ArgumentHelper.GetIntNullable(arguments, "conditionIndex");
        var conditionStr = ArgumentHelper.GetStringNullable(arguments, "condition");
        var value = ArgumentHelper.GetStringNullable(arguments, "value");
        var backgroundColor = ArgumentHelper.GetStringNullable(arguments, "backgroundColor");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");
        }

        var fcs = conditionalFormattings[conditionalFormattingIndex];
        var changes = new List<string>();

        if (conditionIndex.HasValue)
        {
            if (conditionIndex.Value < 0 || conditionIndex.Value >= fcs.Count)
            {
                throw new ArgumentException($"Condition index {conditionIndex.Value} is out of range");
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
                changes.Add($"Operator: {conditionStr}");
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
                changes.Add($"Value: {value}");
            }

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                var style = condition.Style;
                // Handle both color names (e.g., "Red", "Yellow") and hex values (e.g., "#FF0000")
                try
                {
                    style.BackgroundColor = ColorHelper.ParseColor(backgroundColor);
                }
                catch
                {
                    // Fallback to default color if parsing fails
                    style.BackgroundColor = System.Drawing.Color.Yellow;
                }
                style.Pattern = BackgroundType.Solid;
                changes.Add($"Background color: {backgroundColor}");
            }
        }
        
        // Force recalculation to ensure conditional formatting is applied
        workbook.CalculateFormula();

        workbook.Save(outputPath);

        var result = $"Successfully edited conditional formatting #{conditionalFormattingIndex}\n";
        if (changes.Count > 0)
        {
            result += "Changes:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "No changes.\n";
        }
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    /// Deletes conditional formatting from a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing index</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var conditionalFormattingIndex = ArgumentHelper.GetInt(arguments, "conditionalFormattingIndex");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");
        }

        conditionalFormattings.RemoveAt(conditionalFormattingIndex);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);
        
        var remainingCount = conditionalFormattings.Count;
        
        return await Task.FromResult($"Successfully deleted conditional formatting #{conditionalFormattingIndex}\nRemaining conditional formattings in worksheet: {remainingCount}\nOutput: {outputPath}");
    }

    /// <summary>
    /// Gets all conditional formatting rules from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all conditional formatting rules</returns>
    private async Task<string> GetConditionalFormattingAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;
        var result = new StringBuilder();

        result.AppendLine($"=== Conditional formatting information for worksheet '{worksheet.Name}' ===\n");
        result.AppendLine($"Total conditional formattings: {conditionalFormattings.Count}\n");

        if (conditionalFormattings.Count == 0)
        {
            result.AppendLine("No conditional formattings found");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < conditionalFormattings.Count; i++)
        {
            var cf = conditionalFormattings[i];
            result.AppendLine($"[Conditional formatting {i}]");
            result.AppendLine($"Conditional formatting collection index: {i}");
            result.AppendLine("Status: Conditional formatting applied");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

}

