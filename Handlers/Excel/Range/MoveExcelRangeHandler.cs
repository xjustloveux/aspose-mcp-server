using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for moving Excel ranges.
/// </summary>
public class MoveExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "move";

    /// <summary>
    ///     Moves a range to another location.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sourceRange, destCell
    ///     Optional: sheetIndex, sourceSheetIndex, destSheetIndex
    /// </param>
    /// <returns>Success message with move details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var sourceSheetIndex = parameters.GetOptional<int?>("sourceSheetIndex");
        var destSheetIndex = parameters.GetOptional<int?>("destSheetIndex");
        var sourceRange = parameters.GetRequired<string>("sourceRange");
        var destCell = parameters.GetRequired<string>("destCell");

        var workbook = context.Document;
        var srcSheetIdx = sourceSheetIndex ?? sheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = destSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        var sourceRangeObj = ExcelHelper.CreateRange(sourceSheet.Cells, sourceRange, "source range");
        var destRangeObj = ExcelHelper.CreateRange(destSheet.Cells, destCell, "destination cell");

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

        // Clear source range
        for (var i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
        for (var j = sourceRangeObj.FirstColumn;
             j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1;
             j++)
            sourceSheet.Cells[i, j].PutValue("");

        MarkModified(context);

        return Success($"Range {sourceRange} moved to {destCell}.");
    }
}
