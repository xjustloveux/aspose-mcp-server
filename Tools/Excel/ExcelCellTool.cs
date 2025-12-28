using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel cells (write, edit, get, clear)
/// </summary>
public class ExcelCellTool : IAsposeTool
{
    private static readonly Regex CellAddressRegex = new(@"^[A-Za-z]{1,3}\d+$", RegexOptions.Compiled);

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
                description = "Cell reference (e.g., 'A1', 'B2', 'AA100')"
            },
            value = new
            {
                type = "string",
                description =
                    "Value to write (required for write, optional for edit). Supports string, number, boolean, and date formats"
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
            calculateFormula = new
            {
                type = "boolean",
                description = "Calculate formulas before reading value (optional, for get, default: false)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for write/edit/clear operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "cell" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "write" => await WriteCellAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditCellAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetCellAsync(path, sheetIndex, arguments),
            "clear" => await ClearCellAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates the cell address format
    /// </summary>
    /// <param name="cell">Cell address to validate</param>
    /// <exception cref="ArgumentException">Thrown when cell address format is invalid</exception>
    private static void ValidateCellAddress(string cell)
    {
        if (!CellAddressRegex.IsMatch(cell))
            throw new ArgumentException(
                $"Invalid cell address format: '{cell}'. Expected format like 'A1', 'B2', 'AA100'");
    }

    /// <summary>
    ///     Writes a value to a cell
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell address and value</param>
    /// <returns>Success message with cell location</returns>
    private Task<string> WriteCellAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var value = ArgumentHelper.GetString(arguments, "value");

            ValidateCellAddress(cell);

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cellObj = worksheet.Cells[cell];

                ExcelHelper.SetCellValue(cellObj, value);

                workbook.Save(outputPath);
                return $"Cell {cell} written with value '{value}' in sheet {sheetIndex}. Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Edits a cell value
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell address and new value</param>
    /// <returns>Success message with cell location</returns>
    private Task<string> EditCellAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var value = ArgumentHelper.GetStringNullable(arguments, "value");
            var formula = ArgumentHelper.GetStringNullable(arguments, "formula");
            var clearValue = ArgumentHelper.GetBool(arguments, "clearValue", false);

            ValidateCellAddress(cell);

            try
            {
                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cellObj = worksheet.Cells[cell];

                if (clearValue)
                    cellObj.PutValue("");
                else if (!string.IsNullOrEmpty(formula))
                    cellObj.Formula = formula;
                else if (!string.IsNullOrEmpty(value))
                    ExcelHelper.SetCellValue(cellObj, value);
                else
                    throw new ArgumentException("Either value, formula, or clearValue must be provided");

                workbook.Save(outputPath);
                return $"Cell {cell} edited in sheet {sheetIndex}. Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Gets a cell value
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell and optional includeFormat, calculateFormula</param>
    /// <returns>JSON string with cell information</returns>
    private Task<string> GetCellAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var includeFormula = ArgumentHelper.GetBool(arguments, "includeFormula", true);
            var includeFormat = ArgumentHelper.GetBool(arguments, "includeFormat", false);
            var calculateFormula = ArgumentHelper.GetBool(arguments, "calculateFormula", false);

            ValidateCellAddress(cell);

            try
            {
                using var workbook = new Workbook(path);

                if (calculateFormula)
                    workbook.CalculateFormula();

                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var cellObj = worksheet.Cells[cell];

                object? resultObj;

                if (includeFormat)
                {
                    var style = cellObj.GetStyle();
                    resultObj = new
                    {
                        cell,
                        value = cellObj.Value?.ToString() ?? "(empty)",
                        valueType = cellObj.Type.ToString(),
                        formula = includeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null,
                        format = new
                        {
                            fontName = style.Font.Name,
                            fontSize = style.Font.Size,
                            bold = style.Font.IsBold,
                            italic = style.Font.IsItalic,
                            backgroundColor = style.ForegroundColor.ToString(),
                            numberFormat = style.Number
                        }
                    };
                }
                else
                {
                    resultObj = new
                    {
                        cell,
                        value = cellObj.Value?.ToString() ?? "(empty)",
                        valueType = cellObj.Type.ToString(),
                        formula = includeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null
                    };
                }

                return JsonSerializer.Serialize(resultObj, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
            }
        });
    }

    /// <summary>
    ///     Clears a cell (content and/or format)
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell and optional clearContent, clearFormat</param>
    /// <returns>Success message</returns>
    private Task<string> ClearCellAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var clearContent = ArgumentHelper.GetBool(arguments, "clearContent", true);
            var clearFormat = ArgumentHelper.GetBool(arguments, "clearFormat", false);

            ValidateCellAddress(cell);

            try
            {
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

                workbook.Save(outputPath);
                return $"Cell {cell} cleared in sheet {sheetIndex}. Output: {outputPath}";
            }
            catch (CellsException ex)
            {
                throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
            }
        });
    }
}