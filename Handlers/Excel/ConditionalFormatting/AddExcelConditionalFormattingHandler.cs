using System.Drawing;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for adding conditional formatting to Excel worksheets.
/// </summary>
public class AddExcelConditionalFormattingHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds conditional formatting to a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, condition, value
    ///     Optional: sheetIndex, formula2, backgroundColor
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var conditionStr = parameters.GetRequired<string>("condition");
        var value = parameters.GetRequired<string>("value");
        var formula2 = parameters.GetOptional<string?>("formula2");
        var backgroundColor = parameters.GetOptional("backgroundColor", "Yellow");

        ExcelConditionalFormattingHelper.ValidateRange(range);

        try
        {
            var workbook = context.Document;
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

            var operatorType = ExcelConditionalFormattingHelper.ParseOperatorType(conditionStr);
            fc.Operator = operatorType;

            string? warningMessage = null;
            if (!ExcelConditionalFormattingHelper.IsValidCondition(conditionStr))
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

            MarkModified(context);

            return Success($"Conditional formatting added to range {range} ({conditionStr}).{warningMessage ?? ""}");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{range}': {ex.Message}");
        }
    }
}
