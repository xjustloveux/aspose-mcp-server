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
        var getParams = ExtractGetResultParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var cellObj = worksheet.Cells[getParams.Cell];

        if (getParams.CalculateBeforeRead) workbook.CalculateFormula();

        var calculatedValue = cellObj.Value;

        if (!string.IsNullOrEmpty(cellObj.Formula) &&
            (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str))))
        {
            calculatedValue = cellObj.DisplayStringValue;
            if (string.IsNullOrEmpty(calculatedValue?.ToString())) calculatedValue = cellObj.Formula;
        }

        var result = new
        {
            cell = getParams.Cell,
            formula = cellObj.Formula,
            calculatedValue = calculatedValue?.ToString() ?? "(empty)",
            valueType = cellObj.Type.ToString()
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get result parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get result parameters.</returns>
    private static GetResultParameters ExtractGetResultParameters(OperationParameters parameters)
    {
        return new GetResultParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("calculateBeforeRead", true)
        );
    }

    /// <summary>
    ///     Record to hold get formula result parameters.
    /// </summary>
    private record GetResultParameters(string Cell, int SheetIndex, bool CalculateBeforeRead);
}
