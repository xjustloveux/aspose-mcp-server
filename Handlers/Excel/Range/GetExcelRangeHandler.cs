using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for getting data from Excel ranges.
/// </summary>
public class GetExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets data from a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex, includeFormulas, calculateFormulas, includeFormat
    /// </param>
    /// <returns>JSON result with range data.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var includeFormulas = parameters.GetOptional("includeFormulas", false);
        var calculateFormulas = parameters.GetOptional("calculateFormulas", false);
        var includeFormat = parameters.GetOptional("includeFormat", false);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (calculateFormulas)
            workbook.CalculateFormula();

        var cells = worksheet.Cells;
        var cellRange = ExcelHelper.CreateRange(cells, range);

        List<object> cellList = [];
        for (var i = 0; i < cellRange.RowCount; i++)
        for (var j = 0; j < cellRange.ColumnCount; j++)
        {
            var cell = cells[cellRange.FirstRow + i, cellRange.FirstColumn + j];
            var cellRef = CellsHelper.CellIndexToName(cellRange.FirstRow + i, cellRange.FirstColumn + j);

            object? displayValue;
            string? formula = null;

            if (includeFormulas && !string.IsNullOrEmpty(cell.Formula)) formula = cell.Formula;

            if (!string.IsNullOrEmpty(cell.Formula))
            {
                displayValue = cell.Value;
                if (displayValue is CellValueType.IsError)
                    displayValue = cell.DisplayStringValue;
                if (displayValue == null || (displayValue is string str && string.IsNullOrEmpty(str)))
                {
                    displayValue = cell.DisplayStringValue;
                    if (string.IsNullOrEmpty(displayValue?.ToString())) displayValue = cell.Formula;
                }
            }
            else
            {
                displayValue = cell.Value ?? cell.DisplayStringValue;
            }

            if (includeFormat)
            {
                var style = cell.GetStyle();
                cellList.Add(new
                {
                    cell = cellRef,
                    value = displayValue?.ToString() ?? "(empty)",
                    formula,
                    format = new
                    {
                        fontName = style.Font.Name,
                        fontSize = style.Font.Size
                    }
                });
            }
            else
            {
                cellList.Add(new
                {
                    cell = cellRef,
                    value = displayValue?.ToString() ?? "(empty)",
                    formula
                });
            }
        }

        return JsonResult(new
        {
            range,
            rowCount = cellRange.RowCount,
            columnCount = cellRange.ColumnCount,
            count = cellList.Count,
            items = cellList
        });
    }
}
