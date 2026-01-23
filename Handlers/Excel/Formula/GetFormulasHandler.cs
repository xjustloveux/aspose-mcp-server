using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Formula;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for getting formulas from Excel worksheets.
/// </summary>
[ResultType(typeof(GetFormulasResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetFormulasParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, getParams.SheetIndex);
        var cells = worksheet.Cells;

        int startRow, endRow, startCol, endCol;

        if (!string.IsNullOrEmpty(getParams.Range))
        {
            try
            {
                var cellRange = ExcelHelper.CreateRange(cells, getParams.Range);
                startRow = cellRange.FirstRow;
                endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                startCol = cellRange.FirstColumn;
                endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid range format: {getParams.Range}", ex);
            }
        }
        else
        {
            startRow = 0;
            endRow = worksheet.Cells.MaxDataRow;
            startCol = 0;
            endCol = worksheet.Cells.MaxDataColumn;
        }

        List<FormulaInfo> formulaList = [];
        for (var row = startRow; row <= endRow && row <= 10000; row++)
        for (var col = startCol; col <= endCol && col <= 1000; col++)
        {
            var cell = cells[row, col];
            if (!string.IsNullOrEmpty(cell.Formula))
                formulaList.Add(new FormulaInfo
                {
                    Cell = CellsHelper.CellIndexToName(row, col),
                    Formula = cell.Formula,
                    Value = cell.Value?.ToString() ?? "(calculating)"
                });
        }

        if (formulaList.Count == 0)
            return new GetFormulasResult
            {
                Count = 0,
                WorksheetName = worksheet.Name,
                Items = Array.Empty<FormulaInfo>(),
                Message = "No formulas found"
            };

        return new GetFormulasResult
        {
            Count = formulaList.Count,
            WorksheetName = worksheet.Name,
            Items = formulaList
        };
    }

    /// <summary>
    ///     Extracts get formulas parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get formulas parameters.</returns>
    private static GetFormulasParameters ExtractGetFormulasParameters(OperationParameters parameters)
    {
        return new GetFormulasParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("range")
        );
    }

    /// <summary>
    ///     Record to hold get formulas parameters.
    /// </summary>
    private sealed record GetFormulasParameters(int SheetIndex, string? Range);
}
