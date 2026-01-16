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
        var getParams = ExtractGetArrayParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, getParams.SheetIndex);
        var cellObj = worksheet.Cells[getParams.Cell];

        if (!cellObj.IsArrayFormula)
        {
            var notFoundResult = new
            {
                cell = getParams.Cell,
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
            cell = getParams.Cell,
            isArrayFormula = true,
            formula = formula ?? "(empty)",
            arrayRange = arrayRange ?? "Unable to determine"
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get array parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get array parameters.</returns>
    private static GetArrayParameters ExtractGetArrayParameters(OperationParameters parameters)
    {
        return new GetArrayParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold get array formula parameters.
    /// </summary>
    private sealed record GetArrayParameters(string Cell, int SheetIndex);
}
