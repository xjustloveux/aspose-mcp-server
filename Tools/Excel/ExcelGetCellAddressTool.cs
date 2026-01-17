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
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
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
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107 - MCP protocol parameter building
        string operation,
        string? cellAddress,
        int? row,
        int? column)
    {
        return operation.ToLowerInvariant() switch
        {
            "from_a1" => BuildFromA1Parameters(cellAddress),
            "from_index" => BuildFromIndexParameters(row, column),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for converting from A1 notation to row/column index.
    /// </summary>
    /// <param name="cellAddress">The cell address in A1 notation.</param>
    /// <returns>OperationParameters configured for A1 to index conversion.</returns>
    private static OperationParameters BuildFromA1Parameters(string? cellAddress)
    {
        var parameters = new OperationParameters();
        if (cellAddress != null) parameters.Set("cellAddress", cellAddress);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for converting from row/column index to A1 notation.
    /// </summary>
    /// <param name="row">The row index (0-based).</param>
    /// <param name="column">The column index (0-based).</param>
    /// <returns>OperationParameters configured for index to A1 conversion.</returns>
    private static OperationParameters BuildFromIndexParameters(int? row, int? column)
    {
        var parameters = new OperationParameters();
        if (row.HasValue) parameters.Set("row", row.Value);
        if (column.HasValue) parameters.Set("column", column.Value);
        return parameters;
    }
}
