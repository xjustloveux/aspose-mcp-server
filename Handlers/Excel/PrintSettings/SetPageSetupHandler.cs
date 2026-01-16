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
        var setParams = ExtractSetPageSetupParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);
        var pageSetup = worksheet.PageSetup;

        var pageSetupOptions = new PageSetupOptions(
            setParams.Orientation, setParams.PaperSize, setParams.LeftMargin, setParams.RightMargin,
            setParams.TopMargin, setParams.BottomMargin, setParams.Header, setParams.Footer,
            setParams.FitToPage, setParams.FitToPagesWide, setParams.FitToPagesTall);
        var changes = ExcelPrintSettingsHelper.ApplyPageSetup(pageSetup, pageSetupOptions);

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Page setup updated ({changesStr}).");
    }

    /// <summary>
    ///     Extracts set page setup parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set page setup parameters.</returns>
    private static SetPageSetupParameters ExtractSetPageSetupParameters(OperationParameters parameters)
    {
        return new SetPageSetupParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("orientation"),
            parameters.GetOptional<string?>("paperSize"),
            parameters.GetOptional<double?>("leftMargin"),
            parameters.GetOptional<double?>("rightMargin"),
            parameters.GetOptional<double?>("topMargin"),
            parameters.GetOptional<double?>("bottomMargin"),
            parameters.GetOptional<string?>("header"),
            parameters.GetOptional<string?>("footer"),
            parameters.GetOptional<bool?>("fitToPage"),
            parameters.GetOptional<int?>("fitToPagesWide"),
            parameters.GetOptional<int?>("fitToPagesTall")
        );
    }

    /// <summary>
    ///     Record to hold set page setup parameters.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="Orientation">The page orientation.</param>
    /// <param name="PaperSize">The paper size.</param>
    /// <param name="LeftMargin">The left margin in inches.</param>
    /// <param name="RightMargin">The right margin in inches.</param>
    /// <param name="TopMargin">The top margin in inches.</param>
    /// <param name="BottomMargin">The bottom margin in inches.</param>
    /// <param name="Header">The header text.</param>
    /// <param name="Footer">The footer text.</param>
    /// <param name="FitToPage">Whether to enable fit to page.</param>
    /// <param name="FitToPagesWide">The number of pages wide.</param>
    /// <param name="FitToPagesTall">The number of pages tall.</param>
    private sealed record SetPageSetupParameters(
        int SheetIndex,
        string? Orientation,
        string? PaperSize,
        double? LeftMargin,
        double? RightMargin,
        double? TopMargin,
        double? BottomMargin,
        string? Header,
        string? Footer,
        bool? FitToPage,
        int? FitToPagesWide,
        int? FitToPagesTall);
}
