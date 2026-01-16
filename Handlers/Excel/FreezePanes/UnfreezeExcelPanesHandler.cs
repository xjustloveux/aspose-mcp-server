using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Handler for unfreezing panes in Excel worksheets.
/// </summary>
public class UnfreezeExcelPanesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "unfreeze";

    /// <summary>
    ///     Removes freeze panes from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractUnfreezeParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        worksheet.UnFreezePanes();

        MarkModified(context);

        return Success("Unfrozen panes.");
    }

    private static UnfreezeParameters ExtractUnfreezeParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        return new UnfreezeParameters(sheetIndex);
    }

    private record UnfreezeParameters(int SheetIndex);
}
