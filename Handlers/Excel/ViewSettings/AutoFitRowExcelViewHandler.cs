using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class AutoFitRowExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "auto_fit_row";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rowIndex = parameters.GetOptional("rowIndex", 0);
        var startColumn = parameters.GetOptional<int?>("startColumn");
        var endColumn = parameters.GetOptional<int?>("endColumn");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (startColumn.HasValue && endColumn.HasValue)
            worksheet.AutoFitRow(rowIndex, startColumn.Value, endColumn.Value);
        else
            worksheet.AutoFitRow(rowIndex);

        MarkModified(context);
        return Success($"Row {rowIndex} auto-fitted.");
    }
}
