using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for writing data to Excel ranges.
/// </summary>
public class WriteExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "write";

    /// <summary>
    ///     Writes data to a range starting at the specified cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: startCell, data
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>Success message with write details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startCell = parameters.GetRequired<string>("startCell");
        var dataJson = parameters.GetRequired<string>("data");

        var dataArray = ExcelRangeHelper.ParseDataArray(dataJson);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var startCellObj = worksheet.Cells[startCell];
        var startRow = startCellObj.Row;
        var startCol = startCellObj.Column;

        var is2DArrayFormat = dataArray.All(item => item is JsonArray);

        if (is2DArrayFormat && dataArray.Count > 0)
            ExcelRangeHelper.Write2DArrayData(workbook, worksheet, startRow, startCol, dataArray);
        else
            ExcelRangeHelper.WriteObjectArrayData(worksheet, dataArray);

        MarkModified(context);

        return Success($"Data written to range starting at {startCell}.");
    }
}
