using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting row and column headers visibility in Excel worksheets.
/// </summary>
public class SetHeadersExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_headers";

    /// <summary>
    ///     Sets the visibility of row and column headers in the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), visible (default: true)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetHeadersParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.IsRowColumnHeadersVisible = p.Visible;

        MarkModified(context);
        return Success($"RowColumnHeaders visibility set to {(p.Visible ? "visible" : "hidden")}.");
    }

    /// <summary>
    ///     Extracts set headers parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetHeadersParameters record containing all extracted values.</returns>
    private static SetHeadersParameters ExtractSetHeadersParameters(OperationParameters parameters)
    {
        return new SetHeadersParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("visible", true)
        );
    }

    /// <summary>
    ///     Record containing parameters for set headers operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Visible">Whether row and column headers should be visible.</param>
    private record SetHeadersParameters(
        int SheetIndex,
        bool Visible);
}
