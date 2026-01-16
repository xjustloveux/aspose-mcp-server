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
        var p = ExtractClearExcelRangeParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, p.Range);

        if (p is { ClearContent: true, ClearFormat: true })
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
            {
                cells[i, j].PutValue("");
                var defaultStyle = workbook.CreateStyle();
                cells[i, j].SetStyle(defaultStyle);
            }
        }
        else if (p.ClearContent)
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                cells[i, j].PutValue("");
        }
        else if (p.ClearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellRange.ApplyStyle(defaultStyle, new StyleFlag { All = true });
        }

        MarkModified(context);

        return Success($"Range {p.Range} cleared.");
    }

    /// <summary>
    ///     Extracts parameters for ClearExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static ClearExcelRangeParameters ExtractClearExcelRangeParameters(OperationParameters parameters)
    {
        return new ClearExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetOptional("clearContent", true),
            parameters.GetOptional("clearFormat", false)
        );
    }

    /// <summary>
    ///     Parameters for ClearExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="Range">The cell range to clear.</param>
    /// <param name="ClearContent">Whether to clear content.</param>
    /// <param name="ClearFormat">Whether to clear format.</param>
    private sealed record ClearExcelRangeParameters(int SheetIndex, string Range, bool ClearContent, bool ClearFormat);
}
