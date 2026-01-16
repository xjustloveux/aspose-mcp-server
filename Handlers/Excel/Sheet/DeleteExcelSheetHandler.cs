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
        var p = ExtractDeleteExcelSheetParameters(parameters);

        var workbook = context.Document;
        ExcelHelper.ValidateSheetIndex(p.SheetIndex, workbook);

        if (workbook.Worksheets.Count <= 1)
            throw new InvalidOperationException("Cannot delete the last worksheet");

        var sheetName = workbook.Worksheets[p.SheetIndex].Name;
        workbook.Worksheets.RemoveAt(p.SheetIndex);

        MarkModified(context);

        return Success($"Worksheet '{sheetName}' (index {p.SheetIndex}) deleted.");
    }

    private static DeleteExcelSheetParameters ExtractDeleteExcelSheetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");

        return new DeleteExcelSheetParameters(sheetIndex);
    }

    private sealed record DeleteExcelSheetParameters(int SheetIndex);
}
