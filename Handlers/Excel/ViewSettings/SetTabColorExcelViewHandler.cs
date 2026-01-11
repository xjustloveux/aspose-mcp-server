using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetTabColorExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_tab_color";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var color = parameters.GetRequired<string>("color");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        var parsedColor = ColorHelper.ParseColor(color);
        worksheet.TabColor = parsedColor;

        MarkModified(context);
        return Success($"Sheet tab color set to {color}.");
    }
}
