using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SplitWindowExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "split_window";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var splitRow = parameters.GetOptional<int?>("splitRow");
        var splitColumn = parameters.GetOptional<int?>("splitColumn");
        var removeSplit = parameters.GetOptional("removeSplit", false);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (removeSplit)
        {
            worksheet.RemoveSplit();
        }
        else if (splitRow.HasValue || splitColumn.HasValue)
        {
            var row = splitRow ?? 0;
            var col = splitColumn ?? 0;
            worksheet.ActiveCell = CellsHelper.CellIndexToName(row, col);
            worksheet.Split();
        }
        else
        {
            throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
        }

        MarkModified(context);
        return removeSplit
            ? Success($"Window split removed for sheet {sheetIndex}.")
            : Success($"Window split at row {splitRow ?? 0}, column {splitColumn ?? 0} for sheet {sheetIndex}.");
    }
}
