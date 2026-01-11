using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for copying Excel ranges.
/// </summary>
public class CopyExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "copy";

    /// <summary>
    ///     Copies a range to another location.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sourceRange, destCell
    ///     Optional: sheetIndex, sourceSheetIndex, destSheetIndex, copyOptions
    /// </param>
    /// <returns>Success message with copy details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var sourceSheetIndex = parameters.GetOptional<int?>("sourceSheetIndex");
        var destSheetIndex = parameters.GetOptional<int?>("destSheetIndex");
        var sourceRange = parameters.GetRequired<string>("sourceRange");
        var destCell = parameters.GetRequired<string>("destCell");
        var copyOptions = parameters.GetOptional("copyOptions", "All");

        var workbook = context.Document;
        var srcSheetIdx = sourceSheetIndex ?? sheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = destSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        Aspose.Cells.Range sourceRangeObj;
        Aspose.Cells.Range destRangeObj;
        try
        {
            sourceRangeObj = sourceSheet.Cells.CreateRange(sourceRange);
            destRangeObj = destSheet.Cells.CreateRange(destCell);
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Invalid range format. Source range: '{sourceRange}', Destination cell: '{destCell}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
        }

        var pasteType = ExcelRangeHelper.GetPasteType(copyOptions);
        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = pasteType });

        MarkModified(context);

        return Success($"Range {sourceRange} copied to {destCell}.");
    }
}
