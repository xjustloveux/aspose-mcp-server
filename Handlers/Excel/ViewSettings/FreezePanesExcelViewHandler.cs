using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for freezing and unfreezing panes in Excel worksheets.
/// </summary>
public class FreezePanesExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "freeze_panes";

    /// <summary>
    ///     Freezes or unfreezes panes at the specified position.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), freezeRow, freezeColumn, unfreeze (default: false)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when neither freezeRow, freezeColumn, nor unfreeze is provided.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractFreezePanesParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p.Unfreeze)
        {
            worksheet.UnFreezePanes();
        }
        else if (p.FreezeRow.HasValue || p.FreezeColumn.HasValue)
        {
            var row = p.FreezeRow ?? 0;
            var col = p.FreezeColumn ?? 0;
            worksheet.FreezePanes(row, col, row, col);
        }
        else
        {
            throw new ArgumentException("Either freezeRow, freezeColumn, or unfreeze must be provided");
        }

        MarkModified(context);
        return p.Unfreeze
            ? Success($"Panes unfrozen for sheet {p.SheetIndex}.")
            : Success(
                $"Panes frozen at row {p.FreezeRow ?? 0}, column {p.FreezeColumn ?? 0} for sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts freeze panes parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A FreezePanesParameters record containing all extracted values.</returns>
    private static FreezePanesParameters ExtractFreezePanesParameters(OperationParameters parameters)
    {
        return new FreezePanesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<int?>("freezeRow"),
            parameters.GetOptional<int?>("freezeColumn"),
            parameters.GetOptional("unfreeze", false)
        );
    }

    /// <summary>
    ///     Record containing parameters for freeze panes operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="FreezeRow">The row at which to freeze panes.</param>
    /// <param name="FreezeColumn">The column at which to freeze panes.</param>
    /// <param name="Unfreeze">Whether to unfreeze panes.</param>
    private sealed record FreezePanesParameters(
        int SheetIndex,
        int? FreezeRow,
        int? FreezeColumn,
        bool Unfreeze);
}
