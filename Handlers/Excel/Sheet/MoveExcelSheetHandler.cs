using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for moving worksheets to different positions in Excel workbooks.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var targetIndex = parameters.GetOptional<int?>("targetIndex");
        var insertAt = parameters.GetOptional<int?>("insertAt");

        if (!targetIndex.HasValue && !insertAt.HasValue)
            throw new ArgumentException("Either targetIndex or insertAt is required for move operation");

        var finalTargetIndex = targetIndex ?? insertAt!.Value;

        var workbook = context.Document;

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Worksheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (finalTargetIndex < 0 || finalTargetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Target index {finalTargetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (sheetIndex == finalTargetIndex)
            return Success($"Worksheet is already at position {sheetIndex}, no move needed.");

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.MoveTo(finalTargetIndex);

        MarkModified(context);

        return Success($"Worksheet '{sheetName}' moved from position {sheetIndex} to {finalTargetIndex}.");
    }
}
