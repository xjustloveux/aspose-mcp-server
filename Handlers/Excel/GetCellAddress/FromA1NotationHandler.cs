using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.GetCellAddress;

/// <summary>
///     Handler for converting cell address from A1 notation to row/column indices.
/// </summary>
public class FromA1NotationHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "from_a1";

    /// <summary>
    ///     Converts a cell address from A1 notation to row/column indices.
    /// </summary>
    /// <param name="context">The operation context (not used for this operation).</param>
    /// <param name="parameters">
    ///     Required: cellAddress (cell reference like "A1", "B2", "AA100")
    /// </param>
    /// <returns>A message containing the cell address in both A1 notation and row/column index format.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cellAddress = parameters.GetRequired<string>("cellAddress");

        CellsHelper.CellNameToIndex(cellAddress, out var row, out var column);

        ExcelGetCellAddressHelper.ValidateIndexBounds(row, column);

        var a1Notation = CellsHelper.CellIndexToName(row, column);
        return $"{a1Notation} = Row {row}, Column {column}";
    }
}
