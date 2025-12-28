using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel conditional formatting (add, edit, delete, get)
///     Merges: ExcelAddConditionalFormattingTool, ExcelEditConditionalFormattingTool,
///     ExcelDeleteConditionalFormattingTool, ExcelGetConditionalFormattingTool
/// </summary>
public class ExcelConditionalFormattingTool : IAsposeTool
{
    private static readonly Regex RangeRegex = new(@"^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$", RegexOptions.Compiled);

    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel conditional formatting. Supports 4 operations: add, edit, delete, get.

You can add multiple conditional formatting rules to the same range by calling the 'add' operation multiple times. Each rule is independent and will be evaluated separately. To add multiple rules, simply call the 'add' operation multiple times with different conditions for the same range.

Usage examples:
- Add conditional formatting: excel_conditional_formatting(operation='add', path='book.xlsx', range='A1:A10', condition='Between', value='10', formula2='100', backgroundColor='#FF0000')
- Add multiple rules: Call 'add' multiple times with different conditions to create multiple rules for the same range
- Edit conditional formatting: excel_conditional_formatting(operation='edit', path='book.xlsx', conditionalFormattingIndex=0, condition='GreaterThan', value='50')
- Delete conditional formatting: excel_conditional_formatting(operation='delete', path='book.xlsx', conditionalFormattingIndex=0)
- Get conditional formatting: excel_conditional_formatting(operation='get', path='book.xlsx', range='A1:A10')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
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
                description =
                    "Cell range (e.g., 'A1:A10', required for add, optional for get - if not provided, returns all conditional formatting rules)"
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
                description = "Condition value / Formula1 (required for add)"
            },
            formula2 = new
            {
                type = "string",
                description = "Second value for 'Between' condition (optional, use instead of comma-separated value)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color for matching cells (optional)"
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
            "add" => await AddConditionalFormattingAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditConditionalFormattingAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteConditionalFormattingAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetConditionalFormattingAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates the range format (e.g., 'A1:B10')
    /// </summary>
    /// <param name="range">Range string to validate</param>
    /// <exception cref="ArgumentException">Thrown when range format is invalid</exception>
    private static void ValidateRange(string range)
    {
        if (!RangeRegex.IsMatch(range))
            throw new ArgumentException($"Invalid range format: '{range}'. Expected format like 'A1:B10', 'C1:D5'");
    }

    /// <summary>
    ///     Parses condition string to OperatorType
    /// </summary>
    /// <param name="conditionStr">Condition string (e.g., 'GreaterThan', 'LessThan')</param>
    /// <param name="defaultOperator">Default operator if parsing fails</param>
    /// <returns>Parsed OperatorType</returns>
    private static OperatorType ParseOperatorType(string? conditionStr,
        OperatorType defaultOperator = OperatorType.GreaterThan)
    {
        if (string.IsNullOrEmpty(conditionStr))
            return defaultOperator;

        return conditionStr.ToLower() switch
        {
            "greaterthan" => OperatorType.GreaterThan,
            "lessthan" => OperatorType.LessThan,
            "between" => OperatorType.Between,
            "equal" => OperatorType.Equal,
            _ => defaultOperator
        };
    }

    /// <summary>
    ///     Checks if the condition string is a valid operator type
    /// </summary>
    /// <param name="conditionStr">Condition string to check</param>
    /// <returns>True if valid, false otherwise</returns>
    private static bool IsValidCondition(string conditionStr)
    {
        var validConditions = new[] { "greaterthan", "lessthan", "between", "equal" };
        return validConditions.Contains(conditionStr.ToLower());
    }

    /// <summary>
    ///     Adds conditional formatting to a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range, condition, value, optional style properties</param>
    /// <returns>Success message</returns>
    private Task<string> AddConditionalFormattingAsync(string path, string outputPath, int sheetIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var conditionStr = ArgumentHelper.GetString(arguments, "condition");
            var value = ArgumentHelper.GetString(arguments, "value");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var backgroundColor = ArgumentHelper.GetString(arguments, "backgroundColor", "Yellow");

            ValidateRange(range);

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

                var formatIndex = worksheet.ConditionalFormattings.Add();
                var fcs = worksheet.ConditionalFormattings[formatIndex];

                var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
                fcs.AddArea(new CellArea
                {
                    StartRow = cellRange.FirstRow,
                    EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
                    StartColumn = cellRange.FirstColumn,
                    EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
                });

                var conditionIndex = fcs.AddCondition(FormatConditionType.CellValue);
                var fc = fcs[conditionIndex];

                var operatorType = ParseOperatorType(conditionStr);
                fc.Operator = operatorType;

                string? warningMessage = null;
                if (!IsValidCondition(conditionStr))
                    warningMessage =
                        $" Warning: Condition type '{conditionStr}' may not be supported. Valid types are: GreaterThan, LessThan, Between, Equal.";

                fc.Formula1 = value;
                if (operatorType == OperatorType.Between)
                {
                    if (!string.IsNullOrEmpty(formula2))
                    {
                        fc.Formula2 = formula2;
                    }
                    else if (value.Contains(','))
                    {
                        var parts = value.Split(',');
                        if (parts.Length >= 2)
                        {
                            fc.Formula1 = parts[0].Trim();
                            fc.Formula2 = parts[1].Trim();
                        }
                    }
                }

                fc.Style.Pattern = BackgroundType.Solid;
                fc.Style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, Color.Yellow);

                workbook.CalculateFormula();
                workbook.Save(outputPath);

                return
                    $"Conditional formatting added to range {range} ({conditionStr}).{warningMessage ?? ""} Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Edits existing conditional formatting
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing index, optional condition, value, style properties</param>
    /// <returns>Success message</returns>
    private Task<string> EditConditionalFormattingAsync(string path, string outputPath, int sheetIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var conditionalFormattingIndex = ArgumentHelper.GetInt(arguments, "conditionalFormattingIndex");
            var conditionIndex = ArgumentHelper.GetIntNullable(arguments, "conditionIndex");
            var conditionStr = ArgumentHelper.GetStringNullable(arguments, "condition");
            var value = ArgumentHelper.GetStringNullable(arguments, "value");
            var formula2 = ArgumentHelper.GetStringNullable(arguments, "formula2");
            var backgroundColor = ArgumentHelper.GetStringNullable(arguments, "backgroundColor");

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var conditionalFormattings = worksheet.ConditionalFormattings;

                if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
                    throw new ArgumentException(
                        $"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

                var fcs = conditionalFormattings[conditionalFormattingIndex];
                var changes = new List<string>();

                if (conditionIndex.HasValue)
                {
                    if (conditionIndex.Value < 0 || conditionIndex.Value >= fcs.Count)
                        throw new ArgumentException($"Condition index {conditionIndex.Value} is out of range");

                    var condition = fcs[conditionIndex.Value];

                    if (!string.IsNullOrEmpty(conditionStr))
                    {
                        condition.Operator = ParseOperatorType(conditionStr, condition.Operator);
                        changes.Add($"Operator={conditionStr}");
                    }

                    if (!string.IsNullOrEmpty(value))
                    {
                        condition.Formula1 = value;
                        if (condition.Operator == OperatorType.Between)
                        {
                            if (!string.IsNullOrEmpty(formula2))
                            {
                                condition.Formula2 = formula2;
                            }
                            else if (value.Contains(','))
                            {
                                var parts = value.Split(',');
                                if (parts.Length >= 2)
                                {
                                    condition.Formula1 = parts[0].Trim();
                                    condition.Formula2 = parts[1].Trim();
                                }
                            }
                        }

                        changes.Add($"Value={value}");
                    }

                    if (!string.IsNullOrEmpty(backgroundColor))
                    {
                        var style = condition.Style;
                        style.Pattern = BackgroundType.Solid;
                        style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, Color.Yellow);
                        changes.Add($"BackgroundColor={backgroundColor}");
                    }
                }

                workbook.CalculateFormula();
                workbook.Save(outputPath);

                var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
                return
                    $"Edited conditional formatting #{conditionalFormattingIndex} ({changesStr}). Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Deletes conditional formatting from a range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing index</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteConditionalFormattingAsync(string path, string outputPath, int sheetIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var conditionalFormattingIndex = ArgumentHelper.GetInt(arguments, "conditionalFormattingIndex");

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var conditionalFormattings = worksheet.ConditionalFormattings;

                if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
                    throw new ArgumentException(
                        $"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

                conditionalFormattings.RemoveAt(conditionalFormattingIndex);
                workbook.Save(outputPath);

                var remainingCount = conditionalFormattings.Count;
                return
                    $"Deleted conditional formatting #{conditionalFormattingIndex} (remaining: {remainingCount}). Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Gets all conditional formatting rules from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with all conditional formatting rules</returns>
    private Task<string> GetConditionalFormattingAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var conditionalFormattings = worksheet.ConditionalFormattings;

                if (conditionalFormattings.Count == 0)
                {
                    var emptyResult = new
                    {
                        count = 0,
                        sheetIndex,
                        worksheetName = worksheet.Name,
                        items = Array.Empty<object>(),
                        message = "No conditional formattings found"
                    };
                    return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
                }

                var formattingList = new List<object>();
                for (var i = 0; i < conditionalFormattings.Count; i++)
                {
                    var fcs = conditionalFormattings[i];

                    var areasList = new List<string>();
                    for (var k = 0; k < fcs.RangeCount; k++)
                    {
                        var area = fcs.GetCellArea(k);
                        areasList.Add(
                            $"{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}");
                    }

                    var conditionsList = new List<object>();
                    for (var j = 0; j < fcs.Count; j++)
                    {
                        var fc = fcs[j];
                        conditionsList.Add(new
                        {
                            index = j,
                            operatorType = fc.Operator.ToString(),
                            formula1 = fc.Formula1,
                            formula2 = fc.Formula2,
                            foregroundColor = fc.Style?.ForegroundColor.ToString(),
                            backgroundColor = fc.Style?.BackgroundColor.ToString()
                        });
                    }

                    formattingList.Add(new
                    {
                        index = i,
                        areas = areasList,
                        conditionsCount = fcs.Count,
                        conditions = conditionsList
                    });
                }

                var result = new
                {
                    count = conditionalFormattings.Count,
                    sheetIndex,
                    worksheetName = worksheet.Name,
                    items = formattingList
                };

                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed: {ex.Message}");
            }
        });
    }
}