using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel conditional formatting (add, edit, delete, get)
/// </summary>
[McpServerToolType]
public class ExcelConditionalFormattingTool
{
    /// <summary>
    ///     Regex pattern for validating Excel range format (e.g., A1:B10).
    /// </summary>
    private static readonly Regex RangeRegex = new(@"^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$", RegexOptions.Compiled);

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelConditionalFormattingTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelConditionalFormattingTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel conditional formatting operation (add, edit, delete, or get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, or get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range (e.g., 'A1:A10', required for add).</param>
    /// <param name="conditionalFormattingIndex">Conditional formatting index (0-based, required for edit/delete).</param>
    /// <param name="conditionIndex">Condition index within the formatting rule (0-based, optional for edit).</param>
    /// <param name="condition">Condition type: GreaterThan, LessThan, Between, Equal (required for add).</param>
    /// <param name="value">Condition value / Formula1 (required for add).</param>
    /// <param name="formula2">Second value for 'Between' condition (optional).</param>
    /// <param name="backgroundColor">Background color for matching cells.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_conditional_formatting")]
    [Description(@"Manage Excel conditional formatting. Supports 4 operations: add, edit, delete, get.

You can add multiple conditional formatting rules to the same range by calling the 'add' operation multiple times. Each rule is independent and will be evaluated separately. To add multiple rules, simply call the 'add' operation multiple times with different conditions for the same range.

Usage examples:
- Add conditional formatting: excel_conditional_formatting(operation='add', path='book.xlsx', range='A1:A10', condition='Between', value='10', formula2='100', backgroundColor='#FF0000')
- Add multiple rules: Call 'add' multiple times with different conditions to create multiple rules for the same range
- Edit conditional formatting: excel_conditional_formatting(operation='edit', path='book.xlsx', conditionalFormattingIndex=0, condition='GreaterThan', value='50')
- Delete conditional formatting: excel_conditional_formatting(operation='delete', path='book.xlsx', conditionalFormattingIndex=0)
- Get conditional formatting: excel_conditional_formatting(operation='get', path='book.xlsx', range='A1:A10')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range (e.g., 'A1:A10', required for add)")]
        string? range = null,
        [Description("Conditional formatting index (0-based, required for edit/delete)")]
        int conditionalFormattingIndex = 0,
        [Description("Condition index within the formatting rule (0-based, optional for edit)")]
        int? conditionIndex = null,
        [Description("Condition type: GreaterThan, LessThan, Between, Equal (required for add)")]
        string? condition = null,
        [Description("Condition value / Formula1 (required for add)")]
        string? value = null,
        [Description("Second value for 'Between' condition (optional)")]
        string? formula2 = null,
        [Description("Background color for matching cells (default: Yellow)")]
        string backgroundColor = "Yellow")
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddConditionalFormatting(ctx, outputPath, sheetIndex, range, condition, value, formula2,
                backgroundColor),
            "edit" => EditConditionalFormatting(ctx, outputPath, sheetIndex, conditionalFormattingIndex, conditionIndex,
                condition, value, formula2, backgroundColor),
            "delete" => DeleteConditionalFormatting(ctx, outputPath, sheetIndex, conditionalFormattingIndex),
            "get" => GetConditionalFormatting(ctx, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates the range format (e.g., 'A1:B10').
    /// </summary>
    /// <param name="range">The range string to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the range format is invalid.</exception>
    private static void ValidateRange(string range)
    {
        if (!RangeRegex.IsMatch(range))
            throw new ArgumentException($"Invalid range format: '{range}'. Expected format like 'A1:B10', 'C1:D5'");
    }

    /// <summary>
    ///     Parses condition string to OperatorType.
    /// </summary>
    /// <param name="conditionStr">The condition string to parse (e.g., 'GreaterThan', 'LessThan').</param>
    /// <param name="defaultOperator">The default operator to return if conditionStr is null or empty.</param>
    /// <returns>The corresponding OperatorType enum value.</returns>
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
    ///     Checks if the condition string is a valid operator type.
    /// </summary>
    /// <param name="conditionStr">The condition string to check.</param>
    /// <returns>True if the condition string is a valid operator type; otherwise, false.</returns>
    private static bool IsValidCondition(string conditionStr)
    {
        var validConditions = new[] { "greaterthan", "lessthan", "between", "equal" };
        return validConditions.Contains(conditionStr.ToLower());
    }

    /// <summary>
    ///     Adds conditional formatting to a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="range">The cell range to apply formatting (e.g., 'A1:A10').</param>
    /// <param name="conditionStr">The condition type (e.g., 'GreaterThan', 'Between').</param>
    /// <param name="value">The condition value or formula1.</param>
    /// <param name="formula2">The second value for 'Between' condition.</param>
    /// <param name="backgroundColor">The background color for matching cells.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the range/operation is invalid.</exception>
    private static string AddConditionalFormatting(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, string? conditionStr, string? value, string? formula2, string backgroundColor)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for add operation");
        if (string.IsNullOrEmpty(conditionStr))
            throw new ArgumentException("condition is required for add operation");
        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("value is required for add operation");

        ValidateRange(range);

        try
        {
            var workbook = ctx.Document;
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

            ctx.Save(outputPath);

            return
                $"Conditional formatting added to range {range} ({conditionStr}).{warningMessage ?? ""} {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
        }
    }

    /// <summary>
    ///     Edits existing conditional formatting.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="conditionalFormattingIndex">The conditional formatting index (0-based).</param>
    /// <param name="conditionIndex">The condition index within the formatting rule (0-based).</param>
    /// <param name="conditionStr">The condition type to set.</param>
    /// <param name="value">The condition value to set.</param>
    /// <param name="formula2">The second value for 'Between' condition.</param>
    /// <param name="backgroundColor">The background color to set.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the index is out of range or the operation fails.</exception>
    private static string EditConditionalFormatting(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int conditionalFormattingIndex, int? conditionIndex, string? conditionStr, string? value,
        string? formula2, string? backgroundColor)
    {
        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var conditionalFormattings = worksheet.ConditionalFormattings;

            if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
                throw new ArgumentException(
                    $"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

            var fcs = conditionalFormattings[conditionalFormattingIndex];
            List<string> changes = [];

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

            ctx.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
            return
                $"Edited conditional formatting #{conditionalFormattingIndex} ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Deletes conditional formatting from a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="conditionalFormattingIndex">The conditional formatting index (0-based) to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the index is out of range or the operation fails.</exception>
    private static string DeleteConditionalFormatting(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int conditionalFormattingIndex)
    {
        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var conditionalFormattings = worksheet.ConditionalFormattings;

            if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
                throw new ArgumentException(
                    $"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

            conditionalFormattings.RemoveAt(conditionalFormattingIndex);

            ctx.Save(outputPath);

            var remainingCount = conditionalFormattings.Count;
            return
                $"Deleted conditional formatting #{conditionalFormattingIndex} (remaining: {remainingCount}). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Gets all conditional formatting rules from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <returns>A JSON string containing all conditional formatting information.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation fails.</exception>
    private static string GetConditionalFormatting(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        try
        {
            var workbook = ctx.Document;
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

            List<object> formattingList = [];
            for (var i = 0; i < conditionalFormattings.Count; i++)
            {
                var fcs = conditionalFormattings[i];

                List<string> areasList = [];
                for (var k = 0; k < fcs.RangeCount; k++)
                {
                    var area = fcs.GetCellArea(k);
                    areasList.Add(
                        $"{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}");
                }

                List<object> conditionsList = [];
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
    }
}