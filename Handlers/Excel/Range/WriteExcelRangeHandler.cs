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
        var p = ExtractWriteExcelRangeParameters(parameters);

        var dataArray = ExcelRangeHelper.ParseDataArray(p.Data);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var startCellObj = worksheet.Cells[p.StartCell];
        var startRow = startCellObj.Row;
        var startCol = startCellObj.Column;

        var is2DArrayFormat = dataArray.All(item => item is JsonArray);

        if (is2DArrayFormat && dataArray.Count > 0)
            ExcelRangeHelper.Write2DArrayData(workbook, worksheet, startRow, startCol, dataArray);
        else
            ExcelRangeHelper.WriteObjectArrayData(worksheet, dataArray);

        MarkModified(context);

        return Success($"Data written to range starting at {p.StartCell}.");
    }

    /// <summary>
    ///     Extracts parameters for WriteExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static WriteExcelRangeParameters ExtractWriteExcelRangeParameters(OperationParameters parameters)
    {
        return new WriteExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("startCell"),
            parameters.GetRequired<string>("data")
        );
    }

    /// <summary>
    ///     Parameters for WriteExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="StartCell">The starting cell for writing data.</param>
    /// <param name="Data">The data to write as JSON string.</param>
    private record WriteExcelRangeParameters(int SheetIndex, string StartCell, string Data);
}
