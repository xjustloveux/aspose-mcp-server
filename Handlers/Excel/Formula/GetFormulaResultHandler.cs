using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for getting formula results from Excel cells.
/// </summary>
public class GetFormulaResultHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_result";

    /// <summary>
    ///     Gets the calculated result of a formula in a cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell
    ///     Optional: sheetIndex (default: 0), calculateBeforeRead (default: true)
    /// </param>
    /// <returns>JSON string containing formula result information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var calculateBeforeRead = parameters.GetOptional("calculateBeforeRead", true);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (calculateBeforeRead) workbook.CalculateFormula();

        var calculatedValue = cellObj.Value;

        if (!string.IsNullOrEmpty(cellObj.Formula))
            if (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str)))
            {
                calculatedValue = cellObj.DisplayStringValue;
                if (string.IsNullOrEmpty(calculatedValue?.ToString())) calculatedValue = cellObj.Formula;
            }

        var result = new
        {
            cell,
            formula = cellObj.Formula,
            calculatedValue = calculatedValue?.ToString() ?? "(empty)",
            valueType = cellObj.Type.ToString()
        };

        return JsonResult(result);
    }
}
