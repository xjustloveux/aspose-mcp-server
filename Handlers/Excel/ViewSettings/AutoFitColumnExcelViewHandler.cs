using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class AutoFitColumnExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "auto_fit_column";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var columnIndex = parameters.GetOptional("columnIndex", 0);
        var startRow = parameters.GetOptional<int?>("startRow");
        var endRow = parameters.GetOptional<int?>("endRow");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (startRow.HasValue && endRow.HasValue)
            worksheet.AutoFitColumn(columnIndex, startRow.Value, endRow.Value);
        else
            worksheet.AutoFitColumn(columnIndex);

        MarkModified(context);
        return Success($"Column {columnIndex} auto-fitted.");
    }
}
