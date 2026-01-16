using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for splitting or removing window splits in Excel worksheets.
/// </summary>
public class SplitWindowExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "split_window";

    /// <summary>
    ///     Splits the window or removes an existing split.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), splitRow, splitColumn, removeSplit (default: false)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when neither splitRow, splitColumn, nor removeSplit is provided.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSplitWindowParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p.RemoveSplit)
        {
            worksheet.RemoveSplit();
        }
        else if (p.SplitRow.HasValue || p.SplitColumn.HasValue)
        {
            var row = p.SplitRow ?? 0;
            var col = p.SplitColumn ?? 0;
            worksheet.ActiveCell = CellsHelper.CellIndexToName(row, col);
            worksheet.Split();
        }
        else
        {
            throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
        }

        MarkModified(context);
        return p.RemoveSplit
            ? Success($"Window split removed for sheet {p.SheetIndex}.")
            : Success($"Window split at row {p.SplitRow ?? 0}, column {p.SplitColumn ?? 0} for sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts split window parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SplitWindowParameters record containing all extracted values.</returns>
    private static SplitWindowParameters ExtractSplitWindowParameters(OperationParameters parameters)
    {
        return new SplitWindowParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<int?>("splitRow"),
            parameters.GetOptional<int?>("splitColumn"),
            parameters.GetOptional("removeSplit", false)
        );
    }

    /// <summary>
    ///     Record containing parameters for split window operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="SplitRow">The row at which to split the window.</param>
    /// <param name="SplitColumn">The column at which to split the window.</param>
    /// <param name="RemoveSplit">Whether to remove the window split.</param>
    private record SplitWindowParameters(
        int SheetIndex,
        int? SplitRow,
        int? SplitColumn,
        bool RemoveSplit);
}
