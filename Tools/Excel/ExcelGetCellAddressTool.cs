using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Tool for converting between Excel cell address formats (A1 notation and row/column index)
/// </summary>
public class ExcelGetCellAddressTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Convert between cell address formats (A1 notation and row/column index).

Usage examples:
- Convert A1 to index: excel_get_cell_address(cellAddress='A1', convertToIndex=true)
- Convert index to A1: excel_get_cell_address(row=0, column=0, convertToIndex=false)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            cellAddress = new
            {
                type = "string",
                description = "Cell address in A1 notation (e.g., 'A1') or row/column format (e.g., '0,0') (required)"
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cellAddress = ArgumentHelper.GetString(arguments, "cellAddress");
            var convertToIndex = ArgumentHelper.GetBool(arguments, "convertToIndex", false);
            var row = ArgumentHelper.GetIntNullable(arguments, "row");
            var column = ArgumentHelper.GetIntNullable(arguments, "column");

            if (row.HasValue && column.HasValue)
            {
                var a1Notation = CellsHelper.CellIndexToName(row.Value, column.Value);
                return $"Row {row.Value}, Column {column.Value} = {a1Notation}";
            }

            if (convertToIndex)
            {
                // Check if input is in row,column format (e.g., "0,0")
                if (TryParseRowColumnFormat(cellAddress, out var parsedRow, out var parsedCol))
                {
                    var a1Notation = CellsHelper.CellIndexToName(parsedRow, parsedCol);
                    return $"Row {parsedRow}, Column {parsedCol} = {a1Notation}";
                }

                // Try A1 notation
                try
                {
                    CellsHelper.CellNameToIndex(cellAddress, out var rowIndex, out var colIndex);
                    return $"{cellAddress} = Row {rowIndex}, Column {colIndex}";
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Invalid cell address format '{cellAddress}': {ex.Message}");
                    throw new ArgumentException($"Invalid cell address format: {cellAddress}");
                }
            }
            else
            {
                // Check if input is in row,column format (e.g., "0,0")
                if (TryParseRowColumnFormat(cellAddress, out var parsedRow, out var parsedCol))
                {
                    var a1Notation = CellsHelper.CellIndexToName(parsedRow, parsedCol);
                    return $"Valid cell address: {a1Notation} (Row {parsedRow}, Column {parsedCol})";
                }

                // Validate A1 notation
                try
                {
                    CellsHelper.CellNameToIndex(cellAddress, out var rowIndex, out var colIndex);
                    return $"Valid cell address: {cellAddress} (Row {rowIndex}, Column {colIndex})";
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Invalid cell address format '{cellAddress}': {ex.Message}");
                    throw new ArgumentException($"Invalid cell address format: {cellAddress}");
                }
            }
        });
    }

    private bool TryParseRowColumnFormat(string input, out int row, out int column)
    {
        row = 0;
        column = 0;

        // Try to parse "row,column" format (e.g., "0,0", "5,3")
        var parts = input.Split(',');
        if (parts.Length == 2)
            if (int.TryParse(parts[0].Trim(), out row) && int.TryParse(parts[1].Trim(), out column))
                return true;

        return false;
    }
}