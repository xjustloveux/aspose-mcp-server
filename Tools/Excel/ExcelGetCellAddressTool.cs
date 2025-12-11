using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class ExcelGetCellAddressTool : IAsposeTool
{
    public string Description => "Convert between cell address formats (A1 notation and row/column index)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            cellAddress = new
            {
                type = "string",
                description = "Cell address in A1 notation (e.g., 'A1') or row/column format (e.g., '0,0')"
            },
            convertToIndex = new
            {
                type = "boolean",
                description = "Convert to row/column index (optional, default: false, converts to A1 if true)"
            },
            row = new
            {
                type = "number",
                description = "Row index (0-based, optional, used with column)"
            },
            column = new
            {
                type = "number",
                description = "Column index (0-based, optional, used with row)"
            }
        },
        required = new[] { "cellAddress" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var cellAddress = arguments?["cellAddress"]?.GetValue<string>() ?? throw new ArgumentException("cellAddress is required");
        var convertToIndex = arguments?["convertToIndex"]?.GetValue<bool?>() ?? false;
        var row = arguments?["row"]?.GetValue<int?>();
        var column = arguments?["column"]?.GetValue<int?>();

        if (row.HasValue && column.HasValue)
        {
            var a1Notation = CellsHelper.CellIndexToName(row.Value, column.Value);
            return await Task.FromResult($"Row {row.Value}, Column {column.Value} = {a1Notation}");
        }
        else if (convertToIndex)
        {
            // Check if input is in row,column format (e.g., "0,0")
            if (TryParseRowColumnFormat(cellAddress, out int parsedRow, out int parsedCol))
            {
                var a1Notation = CellsHelper.CellIndexToName(parsedRow, parsedCol);
                return await Task.FromResult($"Row {parsedRow}, Column {parsedCol} = {a1Notation}");
            }
            else
            {
                // Try A1 notation
                try
                {
                    int rowIndex, colIndex;
                    CellsHelper.CellNameToIndex(cellAddress, out rowIndex, out colIndex);
                    return await Task.FromResult($"{cellAddress} = Row {rowIndex}, Column {colIndex}");
                }
                catch
                {
                    throw new ArgumentException($"Invalid cell address format: {cellAddress}");
                }
            }
        }
        else
        {
            // Check if input is in row,column format (e.g., "0,0")
            if (TryParseRowColumnFormat(cellAddress, out int parsedRow, out int parsedCol))
            {
                var a1Notation = CellsHelper.CellIndexToName(parsedRow, parsedCol);
                return await Task.FromResult($"Valid cell address: {a1Notation} (Row {parsedRow}, Column {parsedCol})");
            }
            else
            {
                // Validate A1 notation
                try
                {
                    int rowIndex, colIndex;
                    CellsHelper.CellNameToIndex(cellAddress, out rowIndex, out colIndex);
                    return await Task.FromResult($"Valid cell address: {cellAddress} (Row {rowIndex}, Column {colIndex})");
                }
                catch
                {
                    throw new ArgumentException($"Invalid cell address format: {cellAddress}");
                }
            }
        }
    }
    
    private bool TryParseRowColumnFormat(string input, out int row, out int column)
    {
        row = 0;
        column = 0;
        
        // Try to parse "row,column" format (e.g., "0,0", "5,3")
        var parts = input.Split(',');
        if (parts.Length == 2)
        {
            if (int.TryParse(parts[0].Trim(), out row) && int.TryParse(parts[1].Trim(), out column))
            {
                return true;
            }
        }
        
        return false;
    }
}

