using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Handler for setting print titles (repeating rows/columns) in Excel worksheets.
/// </summary>
public class SetPrintTitlesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_print_titles";

    /// <summary>
    ///     Sets print titles (rows/columns to repeat on each page).
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), rows, columns, clearTitles (default: false)
    /// </param>
    /// <returns>Success message with print titles details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var rows = parameters.GetOptional<string?>("rows");
        var columns = parameters.GetOptional<string?>("columns");
        var clearTitles = parameters.GetOptional("clearTitles", false);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearTitles)
        {
            worksheet.PageSetup.PrintTitleRows = "";
            worksheet.PageSetup.PrintTitleColumns = "";
        }
        else
        {
            if (!string.IsNullOrEmpty(rows))
                worksheet.PageSetup.PrintTitleRows = rows;
            if (!string.IsNullOrEmpty(columns))
                worksheet.PageSetup.PrintTitleColumns = columns;
        }

        MarkModified(context);

        return Success($"Print titles updated for sheet {sheetIndex}.");
    }
}
