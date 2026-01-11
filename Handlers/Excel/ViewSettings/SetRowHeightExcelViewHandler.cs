using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetRowHeightExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_row_height";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetOptional("rowIndex", 0);
        var height = parameters.GetOptional("height", 15.0);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.Cells.SetRowHeight(rowIndex, height);

        MarkModified(context);
        return Success($"Row {rowIndex} height set to {height} points.");
    }
}
