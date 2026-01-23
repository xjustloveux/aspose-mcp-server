using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for copying Excel ranges.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    /// <exception cref="ArgumentException">Thrown when range format is invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractCopyExcelRangeParameters(parameters);

        var workbook = context.Document;
        var srcSheetIdx = p.SourceSheetIndex ?? p.SheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = p.DestSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        Aspose.Cells.Range sourceRangeObj;
        Aspose.Cells.Range destRangeObj;
        try
        {
            sourceRangeObj = sourceSheet.Cells.CreateRange(p.SourceRange);
            destRangeObj = destSheet.Cells.CreateRange(p.DestCell);
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Invalid range format. Source range: '{p.SourceRange}', Destination cell: '{p.DestCell}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
        }

        var pasteType = ExcelRangeHelper.GetPasteType(p.CopyOptions);
        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = pasteType });

        MarkModified(context);

        return new SuccessResult { Message = $"Range {p.SourceRange} copied to {p.DestCell}." };
    }

    /// <summary>
    ///     Extracts parameters for CopyExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static CopyExcelRangeParameters ExtractCopyExcelRangeParameters(OperationParameters parameters)
    {
        return new CopyExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<int?>("sourceSheetIndex"),
            parameters.GetOptional<int?>("destSheetIndex"),
            parameters.GetRequired<string>("sourceRange"),
            parameters.GetRequired<string>("destCell"),
            parameters.GetOptional("copyOptions", "All")
        );
    }

    /// <summary>
    ///     Parameters for CopyExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The default sheet index.</param>
    /// <param name="SourceSheetIndex">The source sheet index (optional).</param>
    /// <param name="DestSheetIndex">The destination sheet index (optional).</param>
    /// <param name="SourceRange">The source range to copy.</param>
    /// <param name="DestCell">The destination cell.</param>
    /// <param name="CopyOptions">The copy options.</param>
    private sealed record CopyExcelRangeParameters(
        int SheetIndex,
        int? SourceSheetIndex,
        int? DestSheetIndex,
        string SourceRange,
        string DestCell,
        string CopyOptions);
}
