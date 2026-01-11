using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetColumnWidthExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_column_width";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var columnIndex = parameters.GetOptional("columnIndex", 0);
        var width = parameters.GetOptional("width", 8.43);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.Cells.SetColumnWidth(columnIndex, width);

        MarkModified(context);
        return Success($"Column {columnIndex} width set to {width} characters.");
    }
}
