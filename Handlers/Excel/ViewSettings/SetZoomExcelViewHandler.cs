using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetZoomExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_zoom";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var zoom = parameters.GetRequired<int>("zoom");

        if (zoom < 10 || zoom > 400)
            throw new ArgumentException("Zoom must be between 10 and 400");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.Zoom = zoom;

        MarkModified(context);
        return Success($"Zoom level set to {zoom}% for sheet {sheetIndex}.");
    }
}
