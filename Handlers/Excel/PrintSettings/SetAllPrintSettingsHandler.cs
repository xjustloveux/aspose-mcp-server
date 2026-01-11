using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Handler for setting all print settings at once in Excel worksheets.
/// </summary>
public class SetAllPrintSettingsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_all";

    /// <summary>
    ///     Sets all print settings at once.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), range (print area), rows (print title rows),
    ///     columns (print title columns), orientation, paperSize, leftMargin, rightMargin,
    ///     topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall
    /// </param>
    /// <returns>Success message with all settings details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var printArea = parameters.GetOptional<string?>("range");
        var printTitleRows = parameters.GetOptional<string?>("rows");
        var printTitleColumns = parameters.GetOptional<string?>("columns");
        var orientation = parameters.GetOptional<string?>("orientation");
        var paperSize = parameters.GetOptional<string?>("paperSize");
        var leftMargin = parameters.GetOptional<double?>("leftMargin");
        var rightMargin = parameters.GetOptional<double?>("rightMargin");
        var topMargin = parameters.GetOptional<double?>("topMargin");
        var bottomMargin = parameters.GetOptional<double?>("bottomMargin");
        var header = parameters.GetOptional<string?>("header");
        var footer = parameters.GetOptional<string?>("footer");
        var fitToPage = parameters.GetOptional<bool?>("fitToPage");
        var fitToPagesWide = parameters.GetOptional<int?>("fitToPagesWide");
        var fitToPagesTall = parameters.GetOptional<int?>("fitToPagesTall");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pageSetup = worksheet.PageSetup;

        List<string> changes = [];

        if (!string.IsNullOrEmpty(printArea))
        {
            pageSetup.PrintArea = printArea;
            changes.Add($"printArea={printArea}");
        }

        if (!string.IsNullOrEmpty(printTitleRows))
        {
            pageSetup.PrintTitleRows = printTitleRows;
            changes.Add($"printTitleRows={printTitleRows}");
        }

        if (!string.IsNullOrEmpty(printTitleColumns))
        {
            pageSetup.PrintTitleColumns = printTitleColumns;
            changes.Add($"printTitleColumns={printTitleColumns}");
        }

        var pageSetupChanges = ExcelPrintSettingsHelper.ApplyPageSetup(pageSetup, orientation, paperSize, leftMargin,
            rightMargin, topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall);
        changes.AddRange(pageSetupChanges);

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Print settings updated for sheet {sheetIndex} ({changesStr}).");
    }
}
