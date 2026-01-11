using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for deleting conditional formatting from Excel worksheets.
/// </summary>
public class DeleteExcelConditionalFormattingHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes conditional formatting from a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: conditionalFormattingIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var conditionalFormattingIndex = parameters.GetRequired<int>("conditionalFormattingIndex");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;

        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
            throw new ArgumentException(
                $"Conditional formatting index {conditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

        conditionalFormattings.RemoveAt(conditionalFormattingIndex);

        MarkModified(context);

        return Success(
            $"Deleted conditional formatting #{conditionalFormattingIndex} (remaining: {conditionalFormattings.Count}).");
    }
}
