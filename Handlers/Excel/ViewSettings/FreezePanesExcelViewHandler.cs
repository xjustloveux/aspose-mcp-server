using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class FreezePanesExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "freeze_panes";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var freezeRow = parameters.GetOptional<int?>("freezeRow");
        var freezeColumn = parameters.GetOptional<int?>("freezeColumn");
        var unfreeze = parameters.GetOptional("unfreeze", false);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (unfreeze)
        {
            worksheet.UnFreezePanes();
        }
        else if (freezeRow.HasValue || freezeColumn.HasValue)
        {
            var row = freezeRow ?? 0;
            var col = freezeColumn ?? 0;
            worksheet.FreezePanes(row, col, row, col);
        }
        else
        {
            throw new ArgumentException("Either freezeRow, freezeColumn, or unfreeze must be provided");
        }

        MarkModified(context);
        return unfreeze
            ? Success($"Panes unfrozen for sheet {sheetIndex}.")
            : Success($"Panes frozen at row {freezeRow ?? 0}, column {freezeColumn ?? 0} for sheet {sheetIndex}.");
    }
}
