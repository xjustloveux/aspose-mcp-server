using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for getting array formula information from Excel cells.
/// </summary>
public class GetArrayFormulaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_array";

    /// <summary>
    ///     Gets array formula information for a cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON string containing array formula information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (!cellObj.IsArrayFormula)
        {
            var notFoundResult = new
            {
                cell,
                isArrayFormula = false,
                message = "No array formula found in this cell"
            };
            return JsonResult(notFoundResult);
        }

        var formula = cellObj.Formula;
        string? arrayRange;

        try
        {
            var rangeInfo = cellObj.GetArrayRange();
            var startCellName = CellsHelper.CellIndexToName(rangeInfo.StartRow, rangeInfo.StartColumn);
            var endCellName = CellsHelper.CellIndexToName(rangeInfo.EndRow, rangeInfo.EndColumn);
            arrayRange = $"{startCellName}:{endCellName}";
        }
        catch
        {
            arrayRange = null;
        }

        var result = new
        {
            cell,
            isArrayFormula = true,
            formula = formula ?? "(empty)",
            arrayRange = arrayRange ?? "Unable to determine"
        };

        return JsonResult(result);
    }
}
