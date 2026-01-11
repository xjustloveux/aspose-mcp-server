using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetGridlinesExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_gridlines";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var visible = parameters.GetOptional("visible", true);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.IsGridlinesVisible = visible;

        MarkModified(context);
        return Success($"Gridlines visibility set to {(visible ? "visible" : "hidden")}.");
    }
}
