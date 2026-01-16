using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for auto-fitting column width in Excel worksheets.
/// </summary>
public class AutoFitColumnExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "auto_fit_column";

    /// <summary>
    ///     Auto-fits a column width based on its content.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), columnIndex (default: 0), startRow, endRow
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractAutoFitColumnParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p is { StartRow: { } startRow, EndRow: { } endRow })
            worksheet.AutoFitColumn(p.ColumnIndex, startRow, endRow);
        else
            worksheet.AutoFitColumn(p.ColumnIndex);

        MarkModified(context);
        return Success($"Column {p.ColumnIndex} auto-fitted.");
    }

    /// <summary>
    ///     Extracts auto-fit column parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>An AutoFitColumnParameters record containing all extracted values.</returns>
    private static AutoFitColumnParameters ExtractAutoFitColumnParameters(OperationParameters parameters)
    {
        return new AutoFitColumnParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("columnIndex", 0),
            parameters.GetOptional<int?>("startRow"),
            parameters.GetOptional<int?>("endRow")
        );
    }

    /// <summary>
    ///     Record containing parameters for auto-fit column operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="ColumnIndex">The index of the column to auto-fit.</param>
    /// <param name="StartRow">The starting row for auto-fit range.</param>
    /// <param name="EndRow">The ending row for auto-fit range.</param>
    private sealed record AutoFitColumnParameters(
        int SheetIndex,
        int ColumnIndex,
        int? StartRow,
        int? EndRow);
}
