using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting zoom level in Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetZoomExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_zoom";

    /// <summary>
    ///     Sets the zoom level for a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: zoom (10-400)
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when zoom value is not between 10 and 400.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetZoomParameters(parameters);

        if (p.Zoom < 10 || p.Zoom > 400)
            throw new ArgumentException("Zoom must be between 10 and 400");

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.Zoom = p.Zoom;

        MarkModified(context);
        return new SuccessResult { Message = $"Zoom level set to {p.Zoom}% for sheet {p.SheetIndex}." };
    }

    /// <summary>
    ///     Extracts set zoom parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetZoomParameters record containing all extracted values.</returns>
    private static SetZoomParameters ExtractSetZoomParameters(OperationParameters parameters)
    {
        return new SetZoomParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("zoom")
        );
    }

    /// <summary>
    ///     Record containing parameters for set zoom operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Zoom">The zoom level percentage (10-400).</param>
    private sealed record SetZoomParameters(
        int SheetIndex,
        int Zoom);
}
