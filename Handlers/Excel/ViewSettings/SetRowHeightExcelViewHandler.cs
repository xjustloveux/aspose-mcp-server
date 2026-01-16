using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting row height in Excel worksheets.
/// </summary>
public class SetRowHeightExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_row_height";

    /// <summary>
    ///     Sets the height of a specific row.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), rowIndex (default: 0), height (default: 15.0)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetRowHeightParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.Cells.SetRowHeight(p.RowIndex, p.Height);

        MarkModified(context);
        return Success($"Row {p.RowIndex} height set to {p.Height} points.");
    }

    /// <summary>
    ///     Extracts set row height parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetRowHeightParameters record containing all extracted values.</returns>
    private static SetRowHeightParameters ExtractSetRowHeightParameters(OperationParameters parameters)
    {
        return new SetRowHeightParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("rowIndex", 0),
            parameters.GetOptional("height", 15.0)
        );
    }

    /// <summary>
    ///     Record containing parameters for set row height operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="RowIndex">The index of the row.</param>
    /// <param name="Height">The height of the row in points.</param>
    private sealed record SetRowHeightParameters(
        int SheetIndex,
        int RowIndex,
        double Height);
}
