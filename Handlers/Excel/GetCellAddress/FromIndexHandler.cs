using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.GetCellAddress;

/// <summary>
///     Handler for converting cell address from row/column indices to A1 notation.
/// </summary>
public class FromIndexHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "from_index";

    /// <summary>
    ///     Converts row/column indices to A1 notation.
    /// </summary>
    /// <param name="context">The operation context (not used for this operation).</param>
    /// <param name="parameters">
    ///     Required: row (0-based row index), column (0-based column index)
    /// </param>
    /// <returns>A message containing the cell address in both A1 notation and row/column index format.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractFromIndexParameters(parameters);

        ExcelGetCellAddressHelper.ValidateIndexBounds(p.Row, p.Column);

        var a1Notation = CellsHelper.CellIndexToName(p.Row, p.Column);
        return $"{a1Notation} = Row {p.Row}, Column {p.Column}";
    }

    private static FromIndexParameters ExtractFromIndexParameters(OperationParameters parameters)
    {
        var row = parameters.GetRequired<int>("row");
        var column = parameters.GetRequired<int>("column");

        return new FromIndexParameters(row, column);
    }

    private sealed record FromIndexParameters(int Row, int Column);
}
