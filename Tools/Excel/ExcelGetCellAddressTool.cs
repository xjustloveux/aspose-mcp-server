using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Tool for converting between Excel cell address formats (A1 notation and row/column index).
/// </summary>
[McpServerToolType]
public class ExcelGetCellAddressTool
{
    /// <summary>
    ///     Handler registry for cell address operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelGetCellAddressTool" /> class.
    /// </summary>
    public ExcelGetCellAddressTool()
    {
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.GetCellAddress");
    }

    /// <summary>
    ///     Executes an Excel cell address conversion operation.
    /// </summary>
    /// <param name="operation">The operation to perform: from_a1 or from_index.</param>
    /// <param name="cellAddress">Cell address in A1 notation (e.g., 'A1', 'B2', 'AA100'). Required for from_a1.</param>
    /// <param name="row">Row index (0-based, 0 to 1048575). Required for from_index.</param>
    /// <param name="column">Column index (0-based, 0 to 16383). Required for from_index.</param>
    /// <returns>A message containing the cell address in both A1 notation and row/column index format.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid or out of range.</exception>
    [McpServerTool(Name = "excel_get_cell_address")]
    [Description(@"Convert between cell address formats (A1 notation and row/column index).

Usage examples:
- Convert A1 to index: excel_get_cell_address(operation='from_a1', cellAddress='B2') returns row/column index
- Convert index to A1: excel_get_cell_address(operation='from_index', row=1, column=1) returns 'B2'")]
    public string Execute(
        [Description("Operation: from_a1 (convert A1 to index), from_index (convert index to A1)")]
        string operation,
        [Description("Cell address in A1 notation (e.g., 'A1', 'B2', 'AA100'). Required for from_a1.")]
        string? cellAddress = null,
        [Description("Row index (0-based, 0 to 1048575). Required for from_index.")]
        int? row = null,
        [Description("Column index (0-based, 0 to 16383). Required for from_index.")]
        int? column = null)
    {
        var parameters = BuildParameters(operation, cellAddress, row, column);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = new Workbook()
        };

        return handler.Execute(operationContext, parameters);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        string? cellAddress,
        int? row,
        int? column)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "from_a1":
                if (cellAddress != null) parameters.Set("cellAddress", cellAddress);
                break;

            case "from_index":
                if (row.HasValue) parameters.Set("row", row.Value);
                if (column.HasValue) parameters.Set("column", column.Value);
                break;
        }

        return parameters;
    }
}
