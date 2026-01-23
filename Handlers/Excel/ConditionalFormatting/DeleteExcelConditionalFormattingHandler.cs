using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for deleting conditional formatting from Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;

        if (deleteParams.ConditionalFormattingIndex < 0 ||
            deleteParams.ConditionalFormattingIndex >= conditionalFormattings.Count)
            throw new ArgumentException(
                $"Conditional formatting index {deleteParams.ConditionalFormattingIndex} is out of range (worksheet has {conditionalFormattings.Count} conditional formattings)");

        conditionalFormattings.RemoveAt(deleteParams.ConditionalFormattingIndex);

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Deleted conditional formatting #{deleteParams.ConditionalFormattingIndex} (remaining: {conditionalFormattings.Count})."
        };
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("conditionalFormattingIndex"));
    }

    private sealed record DeleteParameters(int SheetIndex, int ConditionalFormattingIndex);
}
