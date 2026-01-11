using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for adding formulas to Excel cells.
/// </summary>
public class AddFormulaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a formula to a cell in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell, formula
    ///     Optional: sheetIndex (default: 0), autoCalculate (default: true)
    /// </param>
    /// <returns>Success message with formula details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var formula = parameters.GetRequired<string>("formula");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var autoCalculate = parameters.GetOptional("autoCalculate", true);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var cellObj = worksheet.Cells[cell];
        cellObj.Formula = formula;

        string? warningMessage = null;
        if (autoCalculate)
        {
            workbook.CalculateFormula();
            _ = cellObj.Value;

            if (cellObj.Type == CellValueType.IsError)
            {
                var errorValue = cellObj.DisplayStringValue;
                if (!string.IsNullOrEmpty(errorValue) && errorValue.StartsWith("#"))
                {
                    warningMessage = $" Warning: {errorValue}";
                    warningMessage += errorValue switch
                    {
                        "#NAME?" => " (invalid function name)",
                        "#VALUE?" => " (incorrect argument type)",
                        "#REF!" => " (invalid cell reference)",
                        _ => ""
                    };
                }
            }
        }

        MarkModified(context);

        var result = $"Formula added to {cell}: {formula}";
        if (!string.IsNullOrEmpty(warningMessage)) result += $".{warningMessage}";
        return result;
    }
}
