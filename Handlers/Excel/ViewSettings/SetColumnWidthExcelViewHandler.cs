using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting column width in Excel worksheets.
/// </summary>
public class SetColumnWidthExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_column_width";

    /// <summary>
    ///     Sets the width of a specific column.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), columnIndex (default: 0), width (default: 8.43)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetColumnWidthParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.Cells.SetColumnWidth(p.ColumnIndex, p.Width);

        MarkModified(context);
        return Success($"Column {p.ColumnIndex} width set to {p.Width} characters.");
    }

    /// <summary>
    ///     Extracts set column width parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetColumnWidthParameters record containing all extracted values.</returns>
    private static SetColumnWidthParameters ExtractSetColumnWidthParameters(OperationParameters parameters)
    {
        return new SetColumnWidthParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("columnIndex", 0),
            parameters.GetOptional("width", 8.43)
        );
    }

    /// <summary>
    ///     Record containing parameters for set column width operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="ColumnIndex">The index of the column.</param>
    /// <param name="Width">The width of the column in characters.</param>
    private record SetColumnWidthParameters(
        int SheetIndex,
        int ColumnIndex,
        double Width);
}
