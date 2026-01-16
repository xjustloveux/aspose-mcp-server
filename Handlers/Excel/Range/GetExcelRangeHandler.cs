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
        var p = ExtractGetExcelRangeParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        if (p.CalculateFormulas)
            workbook.CalculateFormula();

        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, p.Range);
        var cellList = CollectCellData(worksheet.Cells, cellRange, p.IncludeFormulas, p.IncludeFormat);

        return JsonResult(new
        {
            range = p.Range,
            rowCount = cellRange.RowCount,
            columnCount = cellRange.ColumnCount,
            count = cellList.Count,
            items = cellList
        });
    }

    /// <summary>
    ///     Collects cell data from a range into a list of cell objects.
    /// </summary>
    /// <param name="cells">The cells collection.</param>
    /// <param name="cellRange">The range to collect data from.</param>
    /// <param name="includeFormulas">Whether to include formulas in the output.</param>
    /// <param name="includeFormat">Whether to include format information.</param>
    /// <returns>A list of cell data objects.</returns>
    private static List<object> CollectCellData(Cells cells, Aspose.Cells.Range cellRange, bool includeFormulas,
        bool includeFormat)
    {
        List<object> cellList = [];
        for (var i = 0; i < cellRange.RowCount; i++)
        for (var j = 0; j < cellRange.ColumnCount; j++)
        {
            var row = cellRange.FirstRow + i;
            var col = cellRange.FirstColumn + j;
            var cell = cells[row, col];
            cellList.Add(BuildCellObject(cell, row, col, includeFormulas, includeFormat));
        }

        return cellList;
    }

    /// <summary>
    ///     Builds a cell data object for output.
    /// </summary>
    /// <param name="cell">The cell to get data from.</param>
    /// <param name="row">The row index.</param>
    /// <param name="col">The column index.</param>
    /// <param name="includeFormulas">Whether to include formula information.</param>
    /// <param name="includeFormat">Whether to include format information.</param>
    /// <returns>An anonymous object containing cell data.</returns>
    private static object BuildCellObject(Aspose.Cells.Cell cell, int row, int col, bool includeFormulas,
        bool includeFormat)
    {
        var cellRef = CellsHelper.CellIndexToName(row, col);
        var formula = includeFormulas && !string.IsNullOrEmpty(cell.Formula) ? cell.Formula : null;
        var displayValue = GetDisplayValue(cell);
        var valueStr = displayValue?.ToString() ?? "(empty)";

        if (includeFormat)
        {
            var style = cell.GetStyle();
            return new
            {
                cell = cellRef,
                value = valueStr,
                formula,
                format = new { fontName = style.Font.Name, fontSize = style.Font.Size }
            };
        }

        return new { cell = cellRef, value = valueStr, formula };
    }

    /// <summary>
    ///     Gets the display value of a cell, handling formulas and errors appropriately.
    /// </summary>
    /// <param name="cell">The cell to get the display value from.</param>
    /// <returns>The display value of the cell.</returns>
    private static object? GetDisplayValue(Aspose.Cells.Cell cell)
    {
        if (string.IsNullOrEmpty(cell.Formula))
            return cell.Value ?? cell.DisplayStringValue;

        var displayValue = cell.Value;
        if (displayValue is CellValueType.IsError)
            displayValue = cell.DisplayStringValue;

        if (displayValue != null && !(displayValue is string str && string.IsNullOrEmpty(str)))
            return displayValue;

        displayValue = cell.DisplayStringValue;
        return string.IsNullOrEmpty(displayValue?.ToString()) ? cell.Formula : displayValue;
    }

    /// <summary>
    ///     Extracts parameters for GetExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static GetExcelRangeParameters ExtractGetExcelRangeParameters(OperationParameters parameters)
    {
        return new GetExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetOptional("includeFormulas", false),
            parameters.GetOptional("calculateFormulas", false),
            parameters.GetOptional("includeFormat", false)
        );
    }

    /// <summary>
    ///     Parameters for GetExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="Range">The cell range to get data from.</param>
    /// <param name="IncludeFormulas">Whether to include formulas.</param>
    /// <param name="CalculateFormulas">Whether to calculate formulas before getting data.</param>
    /// <param name="IncludeFormat">Whether to include format information.</param>
    private record GetExcelRangeParameters(
        int SheetIndex,
        string Range,
        bool IncludeFormulas,
        bool CalculateFormulas,
        bool IncludeFormat);
}
