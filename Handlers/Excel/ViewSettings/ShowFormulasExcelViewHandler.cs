using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class ShowFormulasExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "show_formulas";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var visible = parameters.GetOptional("visible", true);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);
        worksheet.ShowFormulas = visible;

        MarkModified(context);
        return Success($"Formulas {(visible ? "shown" : "hidden")} for sheet {sheetIndex}.");
    }
}
