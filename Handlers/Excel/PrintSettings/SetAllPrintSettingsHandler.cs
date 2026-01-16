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

        var pageSetupChanges = ExcelPrintSettingsHelper.ApplyPageSetup(pageSetup, setParams.Orientation,
            setParams.PaperSize,
            setParams.LeftMargin, setParams.RightMargin, setParams.TopMargin, setParams.BottomMargin,
            setParams.Header, setParams.Footer, setParams.FitToPage, setParams.FitToPagesWide,
            setParams.FitToPagesTall);
        changes.AddRange(pageSetupChanges);

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Print settings updated for sheet {setParams.SheetIndex} ({changesStr}).");
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
    private record SetAllPrintSettingsParameters(
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
