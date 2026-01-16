using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting all view settings at once in Excel worksheets.
/// </summary>
public class SetAllExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_all";

    /// <summary>
    ///     Sets all view settings at once.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), zoom (default: 100), showGridlines, showRowColumnHeaders,
    ///     showZeroValues, displayRightToLeft
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when zoom value is not between 10 and 400.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetAllParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p.Zoom != 100)
        {
            if (p.Zoom < 10 || p.Zoom > 400)
                throw new ArgumentException("Zoom must be between 10 and 400");
            worksheet.Zoom = p.Zoom;
        }

        if (p.ShowGridlines.HasValue)
            worksheet.IsGridlinesVisible = p.ShowGridlines.Value;

        if (p.ShowRowColumnHeaders.HasValue)
            worksheet.IsRowColumnHeadersVisible = p.ShowRowColumnHeaders.Value;

        if (p.ShowZeroValues.HasValue)
            worksheet.DisplayZeros = p.ShowZeroValues.Value;

        if (p.DisplayRightToLeft.HasValue)
            worksheet.DisplayRightToLeft = p.DisplayRightToLeft.Value;

        MarkModified(context);
        return Success($"View settings updated for sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts set all parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetAllParameters record containing all extracted values.</returns>
    private static SetAllParameters ExtractSetAllParameters(OperationParameters parameters)
    {
        return new SetAllParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("zoom", 100),
            parameters.GetOptional<bool?>("showGridlines"),
            parameters.GetOptional<bool?>("showRowColumnHeaders"),
            parameters.GetOptional<bool?>("showZeroValues"),
            parameters.GetOptional<bool?>("displayRightToLeft")
        );
    }

    /// <summary>
    ///     Record containing parameters for set all view settings operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Zoom">The zoom level percentage.</param>
    /// <param name="ShowGridlines">Whether to show gridlines.</param>
    /// <param name="ShowRowColumnHeaders">Whether to show row and column headers.</param>
    /// <param name="ShowZeroValues">Whether to show zero values.</param>
    /// <param name="DisplayRightToLeft">Whether to display right to left.</param>
    private sealed record SetAllParameters(
        int SheetIndex,
        int Zoom,
        bool? ShowGridlines,
        bool? ShowRowColumnHeaders,
        bool? ShowZeroValues,
        bool? DisplayRightToLeft);
}
