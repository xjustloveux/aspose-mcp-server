using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for editing data in Excel ranges.
/// </summary>
public class EditExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits data in an existing range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, data
    ///     Optional: sheetIndex, clearRange
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var dataJson = parameters.GetRequired<string>("data");
        var clearRange = parameters.GetOptional("clearRange", false);

        var dataArray = ExcelRangeHelper.ParseDataArray(dataJson);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (clearRange)
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                cells[i, j].PutValue("");

        var startRow = cellRange.FirstRow;
        var startCol = cellRange.FirstColumn;

        for (var i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
                for (var j = 0; j < rowArray.Count; j++)
                {
                    var value = rowArray[j]?.GetValue<string>() ?? "";
                    ExcelHelper.SetCellValue(cells[startRow + i, startCol + j], value);
                }
        }

        MarkModified(context);

        return Success($"Range {range} edited.");
    }
}
