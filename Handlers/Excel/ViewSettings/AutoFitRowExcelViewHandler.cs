using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for auto-fitting row height in Excel worksheets.
/// </summary>
public class AutoFitRowExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "auto_fit_row";

    /// <summary>
    ///     Auto-fits a row height based on its content.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), rowIndex (default: 0), startColumn, endColumn
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractAutoFitRowParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p is { StartColumn: { } startColumn, EndColumn: { } endColumn })
            worksheet.AutoFitRow(p.RowIndex, startColumn, endColumn);
        else
            worksheet.AutoFitRow(p.RowIndex);

        MarkModified(context);
        return Success($"Row {p.RowIndex} auto-fitted.");
    }

    /// <summary>
    ///     Extracts auto-fit row parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>An AutoFitRowParameters record containing all extracted values.</returns>
    private static AutoFitRowParameters ExtractAutoFitRowParameters(OperationParameters parameters)
    {
        return new AutoFitRowParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("rowIndex", 0),
            parameters.GetOptional<int?>("startColumn"),
            parameters.GetOptional<int?>("endColumn")
        );
    }

    /// <summary>
    ///     Record containing parameters for auto-fit row operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="RowIndex">The index of the row to auto-fit.</param>
    /// <param name="StartColumn">The starting column for auto-fit range.</param>
    /// <param name="EndColumn">The ending column for auto-fit range.</param>
    private sealed record AutoFitRowParameters(
        int SheetIndex,
        int RowIndex,
        int? StartColumn,
        int? EndColumn);
}
