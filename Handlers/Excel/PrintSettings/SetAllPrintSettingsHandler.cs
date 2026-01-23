using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Handler for setting all print settings at once in Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var setParams = ExtractSetAllPrintSettingsParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);
        var pageSetup = worksheet.PageSetup;

        List<string> changes = [];

        if (!string.IsNullOrEmpty(setParams.PrintArea))
        {
            pageSetup.PrintArea = setParams.PrintArea;
            changes.Add($"printArea={setParams.PrintArea}");
        }

        if (!string.IsNullOrEmpty(setParams.PrintTitleRows))
        {
            pageSetup.PrintTitleRows = setParams.PrintTitleRows;
            changes.Add($"printTitleRows={setParams.PrintTitleRows}");
        }

        if (!string.IsNullOrEmpty(setParams.PrintTitleColumns))
        {
            pageSetup.PrintTitleColumns = setParams.PrintTitleColumns;
            changes.Add($"printTitleColumns={setParams.PrintTitleColumns}");
        }

        var pageSetupOptions = new PageSetupOptions(
            setParams.Orientation, setParams.PaperSize, setParams.LeftMargin, setParams.RightMargin,
            setParams.TopMargin, setParams.BottomMargin, setParams.Header, setParams.Footer,
            setParams.FitToPage, setParams.FitToPagesWide, setParams.FitToPagesTall);
        var pageSetupChanges = ExcelPrintSettingsHelper.ApplyPageSetup(pageSetup, pageSetupOptions);
        changes.AddRange(pageSetupChanges);

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return new SuccessResult
            { Message = $"Print settings updated for sheet {setParams.SheetIndex} ({changesStr})." };
    }

    /// <summary>
    ///     Extracts set all print settings parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set all print settings parameters.</returns>
    private static SetAllPrintSettingsParameters ExtractSetAllPrintSettingsParameters(OperationParameters parameters)
    {
        return new SetAllPrintSettingsParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("range"),
            parameters.GetOptional<string?>("rows"),
            parameters.GetOptional<string?>("columns"),
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
    ///     Record to hold set all print settings parameters.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="PrintArea">The print area range.</param>
    /// <param name="PrintTitleRows">The print title rows.</param>
    /// <param name="PrintTitleColumns">The print title columns.</param>
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
    private sealed record SetAllPrintSettingsParameters(
        int SheetIndex,
        string? PrintArea,
        string? PrintTitleRows,
        string? PrintTitleColumns,
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
