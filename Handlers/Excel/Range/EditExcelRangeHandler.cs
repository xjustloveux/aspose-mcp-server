using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for editing data in Excel ranges.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractEditExcelRangeParameters(parameters);

        var dataArray = ExcelRangeHelper.ParseDataArray(p.Data);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, p.Range);

        if (p.ClearRange)
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

        return new SuccessResult { Message = $"Range {p.Range} edited." };
    }

    /// <summary>
    ///     Extracts parameters for EditExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static EditExcelRangeParameters ExtractEditExcelRangeParameters(OperationParameters parameters)
    {
        return new EditExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetRequired<string>("data"),
            parameters.GetOptional("clearRange", false)
        );
    }

    /// <summary>
    ///     Parameters for EditExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="Range">The cell range to edit.</param>
    /// <param name="Data">The data to write as JSON string.</param>
    /// <param name="ClearRange">Whether to clear the range before writing.</param>
    private sealed record EditExcelRangeParameters(int SheetIndex, string Range, string Data, bool ClearRange);
}
