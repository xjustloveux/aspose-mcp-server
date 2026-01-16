using System.Drawing;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for editing existing conditional formatting in Excel worksheets.
/// </summary>
public class EditExcelConditionalFormattingHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits existing conditional formatting.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: conditionalFormattingIndex
    ///     Optional: sheetIndex, conditionIndex, condition, value, formula2, backgroundColor
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var editParams = ExtractEditParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
            var fcs = GetFormatConditionCollection(worksheet, editParams.ConditionalFormattingIndex);
            var changes = ApplyConditionChanges(fcs, editParams.ConditionIndex, editParams.Condition, editParams.Value,
                editParams.Formula2, editParams.BackgroundColor);

            workbook.CalculateFormula();
            MarkModified(context);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
            return Success($"Edited conditional formatting #{editParams.ConditionalFormattingIndex} ({changesStr}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("conditionalFormattingIndex", 0),
            parameters.GetOptional<int?>("conditionIndex"),
            parameters.GetOptional<string?>("condition"),
            parameters.GetOptional<string?>("value"),
            parameters.GetOptional<string?>("formula2"),
            parameters.GetOptional<string?>("backgroundColor"));
    }

    /// <summary>
    ///     Gets the format condition collection at the specified index.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the conditional formatting.</param>
    /// <param name="index">The index of the conditional formatting collection.</param>
    /// <returns>The format condition collection at the specified index.</returns>
    /// <exception cref="ArgumentException">Thrown when the index is out of range.</exception>
    private static FormatConditionCollection GetFormatConditionCollection(Worksheet worksheet, int index)
    {
        var conditionalFormattings = worksheet.ConditionalFormattings;
        if (index < 0 || index >= conditionalFormattings.Count)
            throw new ArgumentException(
                $"Conditional formatting index {index} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");
        return conditionalFormattings[index];
    }

    /// <summary>
    ///     Applies changes to a format condition based on the provided parameters.
    /// </summary>
    /// <param name="fcs">The format condition collection.</param>
    /// <param name="conditionIndex">The index of the condition to modify.</param>
    /// <param name="conditionStr">The new condition operator string.</param>
    /// <param name="value">The new formula1 value.</param>
    /// <param name="formula2">The new formula2 value for between conditions.</param>
    /// <param name="backgroundColor">The new background color.</param>
    /// <returns>A list of change descriptions.</returns>
    /// <exception cref="ArgumentException">Thrown when the condition index is out of range.</exception>
    private static List<string> ApplyConditionChanges(FormatConditionCollection fcs, int? conditionIndex,
        string? conditionStr, string? value, string? formula2, string? backgroundColor)
    {
        List<string> changes = [];
        if (!conditionIndex.HasValue) return changes;

        if (conditionIndex.Value < 0 || conditionIndex.Value >= fcs.Count)
            throw new ArgumentException($"Condition index {conditionIndex.Value} is out of range");

        var condition = fcs[conditionIndex.Value];
        ApplyOperatorChange(condition, conditionStr, changes);
        ApplyValueChange(condition, value, formula2, changes);
        ApplyBackgroundColorChange(condition, backgroundColor, changes);
        return changes;
    }

    /// <summary>
    ///     Applies the operator change to a format condition.
    /// </summary>
    /// <param name="condition">The format condition to modify.</param>
    /// <param name="conditionStr">The new condition operator string.</param>
    /// <param name="changes">The list to record the change description.</param>
    private static void ApplyOperatorChange(FormatCondition condition, string? conditionStr, List<string> changes)
    {
        if (string.IsNullOrEmpty(conditionStr)) return;
        condition.Operator = ExcelConditionalFormattingHelper.ParseOperatorType(conditionStr, condition.Operator);
        changes.Add($"Operator={conditionStr}");
    }

    /// <summary>
    ///     Applies the value change to a format condition.
    /// </summary>
    /// <param name="condition">The format condition to modify.</param>
    /// <param name="value">The new formula1 value.</param>
    /// <param name="formula2">The new formula2 value for between conditions.</param>
    /// <param name="changes">The list to record the change description.</param>
    private static void ApplyValueChange(FormatCondition condition, string? value, string? formula2,
        List<string> changes)
    {
        if (string.IsNullOrEmpty(value)) return;
        condition.Formula1 = value;

        if (condition.Operator == OperatorType.Between)
            ApplyBetweenFormulas(condition, value, formula2);

        changes.Add($"Value={value}");
    }

    /// <summary>
    ///     Applies the formula values for a between condition.
    /// </summary>
    /// <param name="condition">The format condition to modify.</param>
    /// <param name="value">The value which may contain comma-separated min and max values.</param>
    /// <param name="formula2">The explicit formula2 value if provided.</param>
    private static void ApplyBetweenFormulas(FormatCondition condition, string value, string? formula2)
    {
        if (!string.IsNullOrEmpty(formula2))
        {
            condition.Formula2 = formula2;
            return;
        }

        if (!value.Contains(',')) return;
        var parts = value.Split(',');
        if (parts.Length < 2) return;
        condition.Formula1 = parts[0].Trim();
        condition.Formula2 = parts[1].Trim();
    }

    /// <summary>
    ///     Applies the background color change to a format condition.
    /// </summary>
    /// <param name="condition">The format condition to modify.</param>
    /// <param name="backgroundColor">The new background color string.</param>
    /// <param name="changes">The list to record the change description.</param>
    private static void ApplyBackgroundColorChange(FormatCondition condition, string? backgroundColor,
        List<string> changes)
    {
        if (string.IsNullOrEmpty(backgroundColor)) return;
        var style = condition.Style;
        style.Pattern = BackgroundType.Solid;
        style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, Color.Yellow);
        changes.Add($"BackgroundColor={backgroundColor}");
    }

    private sealed record EditParameters(
        int SheetIndex,
        int ConditionalFormattingIndex,
        int? ConditionIndex,
        string? Condition,
        string? Value,
        string? Formula2,
        string? BackgroundColor);
}
