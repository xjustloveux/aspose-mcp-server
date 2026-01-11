using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetAllExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_all";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var zoom = parameters.GetOptional("zoom", 100);
        var showGridlines = parameters.GetOptional<bool?>("showGridlines");
        var showRowColumnHeaders = parameters.GetOptional<bool?>("showRowColumnHeaders");
        var showZeroValues = parameters.GetOptional<bool?>("showZeroValues");
        var displayRightToLeft = parameters.GetOptional<bool?>("displayRightToLeft");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (zoom != 100)
        {
            if (zoom < 10 || zoom > 400)
                throw new ArgumentException("Zoom must be between 10 and 400");
            worksheet.Zoom = zoom;
        }

        if (showGridlines.HasValue)
            worksheet.IsGridlinesVisible = showGridlines.Value;

        if (showRowColumnHeaders.HasValue)
            worksheet.IsRowColumnHeadersVisible = showRowColumnHeaders.Value;

        if (showZeroValues.HasValue)
            worksheet.DisplayZeros = showZeroValues.Value;

        if (displayRightToLeft.HasValue)
            worksheet.DisplayRightToLeft = displayRightToLeft.Value;

        MarkModified(context);
        return Success($"View settings updated for sheet {sheetIndex}.");
    }
}
