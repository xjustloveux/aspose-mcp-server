using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel formulas (add, get, get_result, calculate, set_array, get_array).
///     Merges: ExcelAddFormulaTool, ExcelGetFormulaTool, ExcelGetFormulaResultTool,
///     ExcelCalculateFormulaTool, ExcelCalculateAllFormulasTool, ExcelSetArrayFormulaTool, ExcelGetArrayFormulaTool.
/// </summary>
public class ExcelFormulaTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel formulas. Supports 6 operations: add, get, get_result, calculate, set_array, get_array.

Usage examples:
- Add formula: excel_formula(operation='add', path='book.xlsx', cell='A1', formula='=SUM(B1:B10)')
- Get formula: excel_formula(operation='get', path='book.xlsx', cell='A1')
- Get result: excel_formula(operation='get_result', path='book.xlsx', cell='A1')
- Calculate: excel_formula(operation='calculate', path='book.xlsx')
- Set array formula: excel_formula(operation='set_array', path='book.xlsx', range='A1:A10', formula='=B1:B10*2')
- Get array formula: excel_formula(operation='get_array', path='book.xlsx', cell='A1')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a formula to a cell (required params: path, cell, formula)
- 'get': Get formula from a cell (required params: path, cell)
- 'get_result': Get formula result (required params: path, cell)
- 'calculate': Calculate all formulas (required params: path)
- 'set_array': Set array formula (required params: path, range, formula)
- 'get_array': Get array formula (required params: path, cell)",
                @enum = new[] { "add", "get", "get_result", "calculate", "set_array", "get_array" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for add/calculate/set_array operations, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1', required for add/get_result/get_array)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10', optional for get, required for set_array)"
            },
            formula = new
            {
                type = "string",
                description = "Formula (e.g., '=SUM(A1:A10)', required for add/set_array)"
            },
            calculateBeforeRead = new
            {
                type = "boolean",
                description = "Calculate formulas before reading (optional, for get_result, default: true)"
            },
            autoCalculate = new
            {
                type = "boolean",
                description =
                    "Automatically calculate formulas after adding (optional, for add/set_array, default: true). Set to false for batch operations to improve performance."
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddFormulaAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetFormulasAsync(path, sheetIndex, arguments),
            "get_result" => await GetFormulaResultAsync(path, sheetIndex, arguments),
            "calculate" => await CalculateFormulasAsync(path, outputPath, sheetIndex),
            "set_array" => await SetArrayFormulaAsync(path, outputPath, sheetIndex, arguments),
            "get_array" => await GetArrayFormulaAsync(path, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a formula to a cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell and formula.</param>
    /// <returns>Success message with cell reference.</returns>
    private Task<string> AddFormulaAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var formula = ArgumentHelper.GetString(arguments, "formula");
            var autoCalculate = ArgumentHelper.GetBool(arguments, "autoCalculate", true);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var cellObj = worksheet.Cells[cell];
            cellObj.Formula = formula;

            string? warningMessage = null;
            if (autoCalculate)
            {
                workbook.CalculateFormula();
                _ = cellObj.Value;

                if (cellObj.Type == CellValueType.IsError)
                {
                    var errorValue = cellObj.DisplayStringValue;
                    if (!string.IsNullOrEmpty(errorValue) && errorValue.StartsWith("#"))
                    {
                        warningMessage = $" Warning: {errorValue}";
                        warningMessage += errorValue switch
                        {
                            "#NAME?" => " (invalid function name)",
                            "#VALUE?" => " (incorrect argument type)",
                            "#REF!" => " (invalid cell reference)",
                            _ => ""
                        };
                    }
                }
            }

            workbook.Save(outputPath);

            var result = $"Formula added to {cell}: {formula}";
            if (!string.IsNullOrEmpty(warningMessage)) result += $".{warningMessage}";
            result += $". Output: {outputPath}";
            return result;
        });
    }

    /// <summary>
    ///     Gets formula from a cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell.</param>
    /// <returns>JSON string with formula information.</returns>
    private Task<string> GetFormulasAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetStringNullable(arguments, "range");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            int startRow, endRow, startCol, endCol;

            if (!string.IsNullOrEmpty(range))
            {
                try
                {
                    var cellRange = ExcelHelper.CreateRange(cells, range);
                    startRow = cellRange.FirstRow;
                    endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                    startCol = cellRange.FirstColumn;
                    endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;
                }
                catch (Exception ex)
                {
                    throw new ArgumentException($"Invalid range format: {range}", ex);
                }
            }
            else
            {
                startRow = 0;
                endRow = worksheet.Cells.MaxDataRow;
                startCol = 0;
                endCol = worksheet.Cells.MaxDataColumn;
            }

            var formulaList = new List<object>();
            for (var row = startRow; row <= endRow && row <= 10000; row++)
            for (var col = startCol; col <= endCol && col <= 1000; col++)
            {
                var cell = cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                    formulaList.Add(new
                    {
                        cell = CellsHelper.CellIndexToName(row, col),
                        formula = cell.Formula,
                        value = cell.Value?.ToString() ?? "(calculating)"
                    });
            }

            if (formulaList.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    worksheetName = worksheet.Name,
                    items = Array.Empty<object>(),
                    message = "No formulas found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var result = new
            {
                count = formulaList.Count,
                worksheetName = worksheet.Name,
                items = formulaList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets the calculated result of a formula.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell.</param>
    /// <returns>JSON string with formula result value.</returns>
    private Task<string> GetFormulaResultAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var calculateBeforeRead = ArgumentHelper.GetBool(arguments, "calculateBeforeRead", true);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];

            if (calculateBeforeRead) workbook.CalculateFormula();

            var calculatedValue = cellObj.Value;

            if (!string.IsNullOrEmpty(cellObj.Formula))
                if (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str)))
                {
                    calculatedValue = cellObj.DisplayStringValue;
                    if (string.IsNullOrEmpty(calculatedValue?.ToString())) calculatedValue = cellObj.Formula;
                }

            var result = new
            {
                cell,
                formula = cellObj.Formula,
                calculatedValue = calculatedValue?.ToString() ?? "(empty)",
                valueType = cellObj.Type.ToString()
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Calculates all formulas in the workbook.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="_">Worksheet index (unused, calculates entire workbook).</param>
    /// <returns>Success message.</returns>
    private Task<string> CalculateFormulasAsync(string path, string outputPath, int _)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            workbook.CalculateFormula();
            workbook.Save(outputPath);

            return $"Formulas calculated. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets an array formula to a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range and formula.</param>
    /// <returns>Success message with range.</returns>
    private Task<string> SetArrayFormulaAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var formula = ArgumentHelper.GetString(arguments, "formula");
            var autoCalculate = ArgumentHelper.GetBool(arguments, "autoCalculate", true);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);
            var cleanFormula = formula.TrimStart('{').TrimEnd('}');

            if (rangeObj.RowCount <= 0 || rangeObj.ColumnCount <= 0)
                throw new ArgumentException(
                    $"Invalid range dimensions: rows={rangeObj.RowCount}, columns={rangeObj.ColumnCount}");

            if (rangeObj.FirstRow < 0 || rangeObj.FirstColumn < 0)
                throw new ArgumentException(
                    $"Invalid range position: startRow={rangeObj.FirstRow}, startColumn={rangeObj.FirstColumn}");

            var firstCell = worksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

#pragma warning disable CS0618
            var formulaToSet = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;

            for (var i = 0; i < rangeObj.RowCount; i++)
            for (var j = 0; j < rangeObj.ColumnCount; j++)
                worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");

            try
            {
                firstCell.SetArrayFormula(formulaToSet, rangeObj.RowCount, rangeObj.ColumnCount);
                if (autoCalculate) workbook.CalculateFormula();

                if (firstCell.IsArrayFormula)
                {
                    workbook.Save(outputPath);
                    return $"Array formula set in range {range}. Output: {outputPath}";
                }

                workbook.Save(outputPath);
                using var verifyWorkbook = new Workbook(outputPath);
                var verifyCell = verifyWorkbook.Worksheets[sheetIndex].Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

                if (verifyCell.IsArrayFormula)
                    return $"Array formula set in range {range}. Output: {outputPath}";

                throw new InvalidOperationException("SetArrayFormula with 2 parameters did not work");
            }
            catch (Exception ex)
            {
                try
                {
                    using var retryWorkbook = new Workbook(path);
                    var retryWorksheet = retryWorkbook.Worksheets[sheetIndex];
                    var retryRangeObj = ExcelHelper.CreateRange(retryWorksheet.Cells, range);
                    var retryFirstCell = retryWorksheet.Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];

                    for (var i = 0; i < retryRangeObj.RowCount; i++)
                    for (var j = 0; j < retryRangeObj.ColumnCount; j++)
                        retryWorksheet.Cells[retryRangeObj.FirstRow + i, retryRangeObj.FirstColumn + j].PutValue("");

                    var formulaWithoutEquals = cleanFormula.StartsWith("=") ? cleanFormula[1..] : cleanFormula;
                    retryFirstCell.SetArrayFormula(formulaWithoutEquals, retryRangeObj.FirstRow,
                        retryRangeObj.FirstColumn, false, false);

                    if (autoCalculate) retryWorkbook.CalculateFormula();
                    retryWorkbook.Save(outputPath);

                    using var verifyWorkbook = new Workbook(outputPath);
                    var verifyCell =
                        verifyWorkbook.Worksheets[sheetIndex].Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];

                    if (verifyCell.IsArrayFormula)
                        return $"Array formula set in range {range}. Output: {outputPath}";

                    throw new InvalidOperationException("SetArrayFormula with 5 parameters did not work");
                }
                catch (Exception ex2)
                {
                    try
                    {
                        using var fallbackWorkbook = new Workbook(path);
                        var fallbackWorksheet = fallbackWorkbook.Worksheets[sheetIndex];
                        var fallbackRangeObj = ExcelHelper.CreateRange(fallbackWorksheet.Cells, range);

                        var formulaWithEquals = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;
                        for (var i = 0; i < fallbackRangeObj.RowCount; i++)
                        for (var j = 0; j < fallbackRangeObj.ColumnCount; j++)
                        {
                            var cell = fallbackWorksheet.Cells[fallbackRangeObj.FirstRow + i,
                                fallbackRangeObj.FirstColumn + j];
                            cell.Formula = formulaWithEquals;
                        }

                        fallbackWorkbook.Save(outputPath);
                        return $"Formula set to range {range} (not a true array formula). Output: {outputPath}";
                    }
                    catch (Exception ex3)
                    {
                        throw new ArgumentException(
                            $"Failed to set array formula. Range: {range}, Formula: {cleanFormula}. Errors: {ex.Message}, {ex2.Message}, {ex3.Message}",
                            ex);
                    }
                }
            }
#pragma warning restore CS0618
        });
    }

    /// <summary>
    ///     Gets array formula information from a cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell.</param>
    /// <returns>JSON string with array formula information.</returns>
    private Task<string> GetArrayFormulaAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];

            if (!cellObj.IsArrayFormula)
            {
                var notFoundResult = new
                {
                    cell,
                    isArrayFormula = false,
                    message = "No array formula found in this cell"
                };
                return JsonSerializer.Serialize(notFoundResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var formula = cellObj.Formula;
            string? arrayRange;

            try
            {
                var range = cellObj.GetArrayRange();
                var startCellName = CellsHelper.CellIndexToName(range.StartRow, range.StartColumn);
                var endCellName = CellsHelper.CellIndexToName(range.EndRow, range.EndColumn);
                arrayRange = $"{startCellName}:{endCellName}";
            }
            catch
            {
                arrayRange = null;
            }

            var result = new
            {
                cell,
                isArrayFormula = true,
                formula = formula ?? "(empty)",
                arrayRange = arrayRange ?? "Unable to determine"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}