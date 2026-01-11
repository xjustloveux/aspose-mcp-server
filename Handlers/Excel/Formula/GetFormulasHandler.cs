using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for getting formulas from Excel worksheets.
/// </summary>
public class GetFormulasHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all formulas from a range or the entire worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), range
    /// </param>
    /// <returns>JSON string containing formula information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetOptional<string?>("range");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        var cells = worksheet.Cells;

        int startRow, endRow, startCol, endCol;

        if (!string.IsNullOrEmpty(range))
        {
            try
            {
                var cellRange = ExcelHelper.CreateRange(cells, range);
                startRow = cellRange.FirstRow;
                endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                startCol = cellRange.FirstColumn;
                endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid range format: {range}", ex);
            }
        }
        else
        {
            startRow = 0;
            endRow = worksheet.Cells.MaxDataRow;
            startCol = 0;
            endCol = worksheet.Cells.MaxDataColumn;
        }

        List<object> formulaList = [];
        for (var row = startRow; row <= endRow && row <= 10000; row++)
        for (var col = startCol; col <= endCol && col <= 1000; col++)
        {
            var cell = cells[row, col];
            if (!string.IsNullOrEmpty(cell.Formula))
                formulaList.Add(new
                {
                    cell = CellsHelper.CellIndexToName(row, col),
                    formula = cell.Formula,
                    value = cell.Value?.ToString() ?? "(calculating)"
                });
        }

        if (formulaList.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No formulas found"
            };
            return JsonResult(emptyResult);
        }

        var result = new
        {
            count = formulaList.Count,
            worksheetName = worksheet.Name,
            items = formulaList
        };

        return JsonResult(result);
    }
}
