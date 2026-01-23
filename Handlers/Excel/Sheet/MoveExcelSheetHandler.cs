using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for moving worksheets to different positions in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MoveExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "move";

    /// <summary>
    ///     Moves a worksheet to a different position in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based source position)
    ///     Required (one of): targetIndex or insertAt (0-based target position)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractMoveExcelSheetParameters(parameters);

        if (!p.TargetIndex.HasValue && !p.InsertAt.HasValue)
            throw new ArgumentException("Either targetIndex or insertAt is required for move operation");

        var finalTargetIndex = p.TargetIndex ?? p.InsertAt!.Value;

        var workbook = context.Document;

        if (p.SheetIndex < 0 || p.SheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Worksheet index {p.SheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (finalTargetIndex < 0 || finalTargetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Target index {finalTargetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (p.SheetIndex == finalTargetIndex)
            return new SuccessResult { Message = $"Worksheet is already at position {p.SheetIndex}, no move needed." };

        var sheetName = workbook.Worksheets[p.SheetIndex].Name;
        var worksheet = workbook.Worksheets[p.SheetIndex];

        worksheet.MoveTo(finalTargetIndex);

        MarkModified(context);

        return new SuccessResult
            { Message = $"Worksheet '{sheetName}' moved from position {p.SheetIndex} to {finalTargetIndex}." };
    }

    private static MoveExcelSheetParameters ExtractMoveExcelSheetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var targetIndex = parameters.GetOptional<int?>("targetIndex");
        var insertAt = parameters.GetOptional<int?>("insertAt");

        return new MoveExcelSheetParameters(sheetIndex, targetIndex, insertAt);
    }

    private sealed record MoveExcelSheetParameters(int SheetIndex, int? TargetIndex, int? InsertAt);
}
