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
        var addParams = ExtractAddParameters(parameters);

        ExcelConditionalFormattingHelper.ValidateRange(addParams.Range);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);

            var formatIndex = worksheet.ConditionalFormattings.Add();
            var fcs = worksheet.ConditionalFormattings[formatIndex];

            var cellRange = ExcelHelper.CreateRange(worksheet.Cells, addParams.Range);
            fcs.AddArea(new CellArea
            {
                StartRow = cellRange.FirstRow,
                EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
                StartColumn = cellRange.FirstColumn,
                EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
            });

            var conditionIndex = fcs.AddCondition(FormatConditionType.CellValue);
            var fc = fcs[conditionIndex];

            var operatorType = ExcelConditionalFormattingHelper.ParseOperatorType(addParams.Condition);
            fc.Operator = operatorType;

            string? warningMessage = null;
            if (!ExcelConditionalFormattingHelper.IsValidCondition(addParams.Condition))
                warningMessage =
                    $" Warning: Condition type '{addParams.Condition}' may not be supported. Valid types are: GreaterThan, LessThan, Between, Equal.";

            fc.Formula1 = addParams.Value;
            if (operatorType == OperatorType.Between)
            {
                if (!string.IsNullOrEmpty(addParams.Formula2))
                {
                    fc.Formula2 = addParams.Formula2;
                }
                else if (addParams.Value.Contains(','))
                {
                    var parts = addParams.Value.Split(',');
                    if (parts.Length >= 2)
                    {
                        fc.Formula1 = parts[0].Trim();
                        fc.Formula2 = parts[1].Trim();
                    }
                }
            }

            fc.Style.Pattern = BackgroundType.Solid;
            fc.Style.ForegroundColor = ColorHelper.ParseColor(addParams.BackgroundColor, Color.Yellow);

            workbook.CalculateFormula();

            MarkModified(context);

            return Success(
                $"Conditional formatting added to range {addParams.Range} ({addParams.Condition}).{warningMessage ?? ""}");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for range '{addParams.Range}': {ex.Message}");
        }
    }

    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetRequired<string>("condition"),
            parameters.GetRequired<string>("value"),
            parameters.GetOptional<string?>("formula2"),
            parameters.GetOptional("backgroundColor", "Yellow"));
    }

    private sealed record AddParameters(
        int SheetIndex,
        string Range,
        string Condition,
        string Value,
        string? Formula2,
        string BackgroundColor);
}
