using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.MergeCells;

/// <summary>
///     Handler for getting merged cells information in Excel workbooks.
/// </summary>
public class GetExcelMergedCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all merged cell ranges from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (0-based, default: 0)
    /// </param>
    /// <returns>JSON string containing the merged cells information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetMergedCellsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var mergedCells = worksheet.Cells.MergedCells;

        if (mergedCells == null || mergedCells.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No merged cells found"
            };
            return JsonResult(emptyResult);
        }

        List<object> mergedList = [];
        for (var i = 0; i < mergedCells.Count; i++)
        {
            var mergedCellObj = mergedCells[i];

            if (mergedCellObj is CellArea cellArea)
            {
                var startCellName = CellsHelper.CellIndexToName(cellArea.StartRow, cellArea.StartColumn);
                var endCellName = CellsHelper.CellIndexToName(cellArea.EndRow, cellArea.EndColumn);
                var rangeName = $"{startCellName}:{endCellName}";

                var cell = worksheet.Cells[cellArea.StartRow, cellArea.StartColumn];
                var cellValue = cell.Value?.ToString() ?? "(empty)";

                mergedList.Add(new
                {
                    index = i,
                    range = rangeName,
                    startCell = startCellName,
                    endCell = endCellName,
                    rowCount = cellArea.EndRow - cellArea.StartRow + 1,
                    columnCount = cellArea.EndColumn - cellArea.StartColumn + 1,
                    value = cellValue
                });
            }
        }

        var result = new
        {
            count = mergedList.Count,
            worksheetName = worksheet.Name,
            items = mergedList
        };

        return JsonResult(result);
    }

    private static GetMergedCellsParameters ExtractGetMergedCellsParameters(OperationParameters parameters)
    {
        return new GetMergedCellsParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    private sealed record GetMergedCellsParameters(int SheetIndex);
}
