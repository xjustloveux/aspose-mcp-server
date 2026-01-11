using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for setting array formulas in Excel.
/// </summary>
public class SetArrayFormulaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_array";

    /// <summary>
    ///     Sets an array formula for a range of cells.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, formula
    ///     Optional: sheetIndex (default: 0), autoCalculate (default: true)
    /// </param>
    /// <returns>Success message with range details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var range = parameters.GetRequired<string>("range");
        var formula = parameters.GetRequired<string>("formula");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var autoCalculate = parameters.GetOptional("autoCalculate", true);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);
        var cleanFormula = formula.TrimStart('{').TrimEnd('}');

        if (rangeObj.RowCount <= 0 || rangeObj.ColumnCount <= 0)
            throw new ArgumentException(
                $"Invalid range dimensions: rows={rangeObj.RowCount}, columns={rangeObj.ColumnCount}");

        if (rangeObj.FirstRow < 0 || rangeObj.FirstColumn < 0)
            throw new ArgumentException(
                $"Invalid range position: startRow={rangeObj.FirstRow}, startColumn={rangeObj.FirstColumn}");

        var firstCell = worksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

#pragma warning disable CS0618
        var formulaToSet = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;

        for (var i = 0; i < rangeObj.RowCount; i++)
        for (var j = 0; j < rangeObj.ColumnCount; j++)
            worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");

        try
        {
            firstCell.SetArrayFormula(formulaToSet, rangeObj.RowCount, rangeObj.ColumnCount);
            if (autoCalculate) workbook.CalculateFormula();

            MarkModified(context);
            return Success($"Array formula set in range {range}.");
        }
        catch (Exception ex)
        {
            try
            {
                for (var i = 0; i < rangeObj.RowCount; i++)
                for (var j = 0; j < rangeObj.ColumnCount; j++)
                    worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");

                var formulaWithoutEquals = cleanFormula.StartsWith("=") ? cleanFormula[1..] : cleanFormula;
                firstCell.SetArrayFormula(formulaWithoutEquals, rangeObj.FirstRow,
                    rangeObj.FirstColumn, false, false);

                if (autoCalculate) workbook.CalculateFormula();

                if (firstCell.IsArrayFormula)
                {
                    MarkModified(context);
                    return Success($"Array formula set in range {range}.");
                }

                throw new InvalidOperationException("SetArrayFormula with 5 parameters did not work");
            }
            catch (Exception ex2)
            {
                try
                {
                    var formulaWithEquals = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;
                    for (var i = 0; i < rangeObj.RowCount; i++)
                    for (var j = 0; j < rangeObj.ColumnCount; j++)
                    {
                        var cell = worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j];
                        cell.Formula = formulaWithEquals;
                    }

                    MarkModified(context);
                    return Success($"Formula set to range {range} (not a true array formula).");
                }
                catch (Exception ex3)
                {
                    throw new ArgumentException(
                        $"Failed to set array formula. Range: {range}, Formula: {cleanFormula}. Errors: {ex.Message}, {ex2.Message}, {ex3.Message}",
                        ex);
                }
            }
        }
#pragma warning restore CS0618
    }
}
