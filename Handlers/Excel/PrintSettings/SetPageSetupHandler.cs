using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Handler for setting page setup options in Excel worksheets.
/// </summary>
public class SetPageSetupHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_page_setup";

    /// <summary>
    ///     Sets page setup options.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), orientation, paperSize, leftMargin, rightMargin,
    ///     topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall
    /// </param>
    /// <returns>Success message with page setup details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
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

        var changes = ExcelPrintSettingsHelper.ApplyPageSetup(pageSetup, orientation, paperSize, leftMargin,
            rightMargin, topMargin, bottomMargin, header, footer, fitToPage, fitToPagesWide, fitToPagesTall);

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Page setup updated ({changesStr}).");
    }
}
