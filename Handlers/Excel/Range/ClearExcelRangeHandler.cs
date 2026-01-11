using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for clearing content and/or format from Excel ranges.
/// </summary>
public class ClearExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears content and/or format from a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex, clearContent, clearFormat
    /// </param>
    /// <returns>Success message with clear details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var clearContent = parameters.GetOptional("clearContent", true);
        var clearFormat = parameters.GetOptional("clearFormat", false);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (clearContent && clearFormat)
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
            {
                cells[i, j].PutValue("");
                var defaultStyle = workbook.CreateStyle();
                cells[i, j].SetStyle(defaultStyle);
            }
        }
        else if (clearContent)
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                cells[i, j].PutValue("");
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellRange.ApplyStyle(defaultStyle, new StyleFlag { All = true });
        }

        MarkModified(context);

        return Success($"Range {range} cleared.");
    }
}
