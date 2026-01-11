using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetHeadersExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_headers";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var visible = parameters.GetOptional("visible", true);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.IsRowColumnHeadersVisible = visible;

        MarkModified(context);
        return Success($"RowColumnHeaders visibility set to {(visible ? "visible" : "hidden")}.");
    }
}
