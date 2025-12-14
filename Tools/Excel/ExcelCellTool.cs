using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel cells (write, edit, get, clear)
/// Merges: ExcelWriteCellTool, ExcelEditCellTool, ExcelGetCellValueTool, ExcelClearCellTool
/// </summary>
public class ExcelCellTool : IAsposeTool
{
    public string Description => @"Manage Excel cells. Supports 4 operations: write, edit, get, clear.

Usage examples:
- Write cell: excel_cell(operation='write', path='book.xlsx', cell='A1', value='Hello')
- Edit cell: excel_cell(operation='edit', path='book.xlsx', cell='A1', value='Updated')
- Get cell: excel_cell(operation='get', path='book.xlsx', cell='A1')
- Clear cell: excel_cell(operation='clear', path='book.xlsx', cell='A1')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'write': Write value to cell (required params: path, cell, value)
- 'edit': Edit cell value (required params: path, cell, value)
- 'get': Get cell value (required params: path, cell)
- 'clear': Clear cell (required params: path, cell)",
                @enum = new[] { "write", "edit", "get", "clear" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1')"
            },
            value = new
            {
                type = "string",
                description = "Value to write (required for write, optional for edit)"
            },
            formula = new
            {
                type = "string",
                description = "Formula to set (optional, for edit, overrides value)"
            },
            clearValue = new
            {
                type = "boolean",
                description = "Clear cell value (optional, for edit)"
            },
            includeFormula = new
            {
                type = "boolean",
                description = "Include formula if present (optional, for get, default: true)"
            },
            includeFormat = new
            {
                type = "boolean",
                description = "Include format information (optional, for get, default: false)"
            },
            clearContent = new
            {
                type = "boolean",
                description = "Clear cell content (optional, for clear, default: true)"
            },
            clearFormat = new
            {
                type = "boolean",
                description = "Clear cell format (optional, for clear, default: false)"
            }
        },
        required = new[] { "operation", "path", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "write" => await WriteCellAsync(arguments, path, sheetIndex),
            "edit" => await EditCellAsync(arguments, path, sheetIndex),
            "get" => await GetCellAsync(arguments, path, sheetIndex),
            "clear" => await ClearCellAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Writes a value to a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and value</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with cell reference</returns>
    private async Task<string> WriteCellAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var value = ArgumentHelper.GetString(arguments, "value", "value");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];
        
        // Parse value as number, boolean, or string
        // Ensures formulas can correctly identify numeric values
        if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numValue))
        {
            cellObj.PutValue(numValue);
        }
        else if (bool.TryParse(value, out bool boolValue))
        {
            cellObj.PutValue(boolValue);
        }
        else
        {
            cellObj.PutValue(value);
        }
        
        workbook.Save(path);

        return await Task.FromResult($"Cell {cell} updated in sheet {sheetIndex}: {path}");
    }

    /// <summary>
    /// Edits a cell value
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and value</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with cell reference</returns>
    private async Task<string> EditCellAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var value = arguments?["value"]?.GetValue<string>();
        var formula = arguments?["formula"]?.GetValue<string>();
        var clearValue = arguments?["clearValue"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (clearValue)
        {
            cellObj.PutValue("");
        }
        else if (!string.IsNullOrEmpty(formula))
        {
            cellObj.Formula = formula;
        }
        else if (!string.IsNullOrEmpty(value))
        {
            // Parse value as number, boolean, or string
            // Ensures formulas can correctly identify numeric values
            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numValue))
            {
                cellObj.PutValue(numValue);
            }
            else if (bool.TryParse(value, out bool boolValue))
            {
                cellObj.PutValue(boolValue);
            }
            else
            {
                cellObj.PutValue(value);
            }
        }
        else
        {
            throw new ArgumentException("Either value, formula, or clearValue must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult($"Cell {cell} edited in sheet {sheetIndex}: {path}");
    }

    /// <summary>
    /// Gets a cell value
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and optional includeFormat</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Cell value and optional format information</returns>
    private async Task<string> GetCellAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var includeFormula = arguments?["includeFormula"]?.GetValue<bool?>() ?? true;
        var includeFormat = arguments?["includeFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        var result = new StringBuilder();
        result.AppendLine($"Cell: {cell}");
        result.AppendLine($"Value: {cellObj.Value ?? "(empty)"}");
        result.AppendLine($"Value Type: {cellObj.Type}");

        if (includeFormula && !string.IsNullOrEmpty(cellObj.Formula))
        {
            result.AppendLine($"Formula: {cellObj.Formula}");
        }

        if (includeFormat)
        {
            var style = cellObj.GetStyle();
            result.AppendLine($"Format:");
            result.AppendLine($"  Font: {style.Font.Name}, Size: {style.Font.Size}");
            result.AppendLine($"  Bold: {style.Font.IsBold}, Italic: {style.Font.IsItalic}");
            result.AppendLine($"  Background Color: {style.ForegroundColor}");
            result.AppendLine($"  Number Format: {style.Number}");
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    /// Clears a cell (content and/or format)
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and optional clearContent, clearFormat</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> ClearCellAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var clearContent = arguments?["clearContent"]?.GetValue<bool?>() ?? true;
        var clearFormat = arguments?["clearFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (clearContent && clearFormat)
        {
            cellObj.PutValue("");
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }
        else if (clearContent)
        {
            cellObj.PutValue("");
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellObj.SetStyle(defaultStyle);
        }

        workbook.Save(path);
        return await Task.FromResult($"Cell {cell} cleared in sheet {sheetIndex}: {path}");
    }
}

