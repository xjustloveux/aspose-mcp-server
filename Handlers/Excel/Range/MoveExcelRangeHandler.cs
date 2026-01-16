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
        var p = ExtractMoveExcelRangeParameters(parameters);

        var workbook = context.Document;
        var srcSheetIdx = p.SourceSheetIndex ?? p.SheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = p.DestSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        var sourceRangeObj = ExcelHelper.CreateRange(sourceSheet.Cells, p.SourceRange, "source range");
        var destRangeObj = ExcelHelper.CreateRange(destSheet.Cells, p.DestCell, "destination cell");

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

        for (var i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
        for (var j = sourceRangeObj.FirstColumn;
             j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1;
             j++)
            sourceSheet.Cells[i, j].PutValue("");

        MarkModified(context);

        return Success($"Range {p.SourceRange} moved to {p.DestCell}.");
    }

    /// <summary>
    ///     Extracts parameters for MoveExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static MoveExcelRangeParameters ExtractMoveExcelRangeParameters(OperationParameters parameters)
    {
        return new MoveExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<int?>("sourceSheetIndex"),
            parameters.GetOptional<int?>("destSheetIndex"),
            parameters.GetRequired<string>("sourceRange"),
            parameters.GetRequired<string>("destCell")
        );
    }

    /// <summary>
    ///     Parameters for MoveExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The default sheet index.</param>
    /// <param name="SourceSheetIndex">The source sheet index (optional).</param>
    /// <param name="DestSheetIndex">The destination sheet index (optional).</param>
    /// <param name="SourceRange">The source range to move.</param>
    /// <param name="DestCell">The destination cell.</param>
    private record MoveExcelRangeParameters(
        int SheetIndex,
        int? SourceSheetIndex,
        int? DestSheetIndex,
        string SourceRange,
        string DestCell);
}
