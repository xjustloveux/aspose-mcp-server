using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting gridlines visibility in Excel worksheets.
/// </summary>
public class SetGridlinesExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_gridlines";

    /// <summary>
    ///     Sets the visibility of gridlines in the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), visible (default: true)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetGridlinesParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.IsGridlinesVisible = p.Visible;

        MarkModified(context);
        return Success($"Gridlines visibility set to {(p.Visible ? "visible" : "hidden")}.");
    }

    /// <summary>
    ///     Extracts set gridlines parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetGridlinesParameters record containing all extracted values.</returns>
    private static SetGridlinesParameters ExtractSetGridlinesParameters(OperationParameters parameters)
    {
        return new SetGridlinesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("visible", true)
        );
    }

    /// <summary>
    ///     Record containing parameters for set gridlines operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Visible">Whether gridlines should be visible.</param>
    private sealed record SetGridlinesParameters(
        int SheetIndex,
        bool Visible);
}
