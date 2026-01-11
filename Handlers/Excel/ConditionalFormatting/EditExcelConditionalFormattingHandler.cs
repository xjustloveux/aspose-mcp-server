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
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var conditionalFormattingIndex = parameters.GetOptional("conditionalFormattingIndex", 0);
        var conditionIndex = parameters.GetOptional<int?>("conditionIndex");
        var conditionStr = parameters.GetOptional<string?>("condition");
        var value = parameters.GetOptional<string?>("value");
        var formula2 = parameters.GetOptional<string?>("formula2");
        var backgroundColor = parameters.GetOptional<string?>("backgroundColor");

        try
        {
            var workbook = context.Document;
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
                    condition.Operator =
                        ExcelConditionalFormattingHelper.ParseOperatorType(conditionStr, condition.Operator);
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

            MarkModified(context);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
            return Success($"Edited conditional formatting #{conditionalFormattingIndex} ({changesStr}).");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }
}
