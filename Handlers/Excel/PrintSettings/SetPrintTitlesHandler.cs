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
        var setParams = ExtractSetPrintTitlesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);

        if (setParams.ClearTitles)
        {
            worksheet.PageSetup.PrintTitleRows = "";
            worksheet.PageSetup.PrintTitleColumns = "";
        }
        else
        {
            if (!string.IsNullOrEmpty(setParams.Rows))
                worksheet.PageSetup.PrintTitleRows = setParams.Rows;
            if (!string.IsNullOrEmpty(setParams.Columns))
                worksheet.PageSetup.PrintTitleColumns = setParams.Columns;
        }

        MarkModified(context);

        return Success($"Print titles updated for sheet {setParams.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts set print titles parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set print titles parameters.</returns>
    private static SetPrintTitlesParameters ExtractSetPrintTitlesParameters(OperationParameters parameters)
    {
        return new SetPrintTitlesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("rows"),
            parameters.GetOptional<string?>("columns"),
            parameters.GetOptional("clearTitles", false)
        );
    }

    /// <summary>
    ///     Record to hold set print titles parameters.
    /// </summary>
    private sealed record SetPrintTitlesParameters(int SheetIndex, string? Rows, string? Columns, bool ClearTitles);
}
