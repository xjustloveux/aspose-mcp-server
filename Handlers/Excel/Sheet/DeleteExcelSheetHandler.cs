using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for deleting worksheets from Excel workbooks.
/// </summary>
public class DeleteExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a worksheet from the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based index of sheet to delete)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");

        var workbook = context.Document;
        ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

        if (workbook.Worksheets.Count <= 1)
            throw new InvalidOperationException("Cannot delete the last worksheet");

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        workbook.Worksheets.RemoveAt(sheetIndex);

        MarkModified(context);

        return Success($"Worksheet '{sheetName}' (index {sheetIndex}) deleted.");
    }
}
