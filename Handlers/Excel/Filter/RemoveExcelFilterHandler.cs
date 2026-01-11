using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for removing auto filter from Excel worksheets.
/// </summary>
public class RemoveExcelFilterHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes auto filter from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.RemoveAutoFilter();

        MarkModified(context);

        return Success($"Auto filter removed from sheet {sheetIndex}.");
    }
}
