using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel formulas (add, get, get_result, calculate, set_array, get_array)
///     Merges: ExcelAddFormulaTool, ExcelGetFormulaTool, ExcelGetFormulaResultTool,
///     ExcelCalculateFormulaTool, ExcelCalculateAllFormulasTool, ExcelSetArrayFormulaTool, ExcelGetArrayFormulaTool
/// </summary>
public class ExcelFormulaTool : IAsposeTool
{
    public string Description =>
        @"Manage Excel formulas. Supports 6 operations: add, get, get_result, calculate, set_array, get_array.

Usage examples:
- Add formula: excel_formula(operation='add', path='book.xlsx', cell='A1', formula='=SUM(B1:B10)')
- Get formula: excel_formula(operation='get', path='book.xlsx', cell='A1')
- Get result: excel_formula(operation='get_result', path='book.xlsx', cell='A1')
- Calculate: excel_formula(operation='calculate', path='book.xlsx')
- Set array formula: excel_formula(operation='set_array', path='book.xlsx', range='A1:A10', formula='=B1:B10*2')
- Get array formula: excel_formula(operation='get_array', path='book.xlsx', cell='A1')";

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
            }
        },
        required = new[] { "operation", "path" }
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
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddFormulaAsync(arguments, path, sheetIndex),
            "get" => await GetFormulasAsync(arguments, path, sheetIndex),
            "get_result" => await GetFormulaResultAsync(arguments, path, sheetIndex),
            "calculate" => await CalculateFormulasAsync(arguments, path, sheetIndex),
            "set_array" => await SetArrayFormulaAsync(arguments, path, sheetIndex),
            "get_array" => await GetArrayFormulaAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a formula to a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and formula</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with cell reference</returns>
    private async Task<string> AddFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell");
        var formula = ArgumentHelper.GetString(arguments, "formula");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var cellObj = worksheet.Cells[cell];
        cellObj.Formula = formula;

        // Calculate formula to ensure correct result
        workbook.CalculateFormula();

        // Ensure calculated result is saved (accessing the value triggers calculation)
        // This ensures that when Excel opens the file, formulas already have correct calculated results
        _ = cellObj.Value;

        // Check for formula calculation errors after calculation
        string? warningMessage = null;
        if (cellObj.Type == CellValueType.IsError)
        {
            var errorValue = cellObj.DisplayStringValue;
            // Common Excel error values: #NAME?, #VALUE?, #REF!, #DIV/0!, #NUM!, #NULL!, #N/A
            if (!string.IsNullOrEmpty(errorValue) && errorValue.StartsWith("#"))
            {
                warningMessage = $"\n⚠️ Warning: Formula calculation resulted in error: {errorValue}";
                if (errorValue == "#NAME?")
                    warningMessage += " This usually indicates an invalid function name or undefined name.";
                else if (errorValue == "#VALUE?")
                    warningMessage += " This usually indicates an incorrect argument type.";
                else if (errorValue == "#REF!") warningMessage += " This usually indicates an invalid cell reference.";
            }
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var result = $"Formula added to cell {cell}: {formula}\nOutput: {outputPath}";
        if (!string.IsNullOrEmpty(warningMessage)) result += warningMessage;
        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Gets formula from a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formula string</returns>
    private async Task<string> GetFormulasAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetStringNullable(arguments, "range");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var result = new StringBuilder();

        result.AppendLine($"=== Formula Information for Worksheet '{worksheet.Name}' ===\n");

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

        var formulaCount = 0;
        for (var row = startRow; row <= endRow && row <= 10000; row++)
        for (var col = startCol; col <= endCol && col <= 1000; col++)
        {
            var cell = cells[row, col];
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                formulaCount++;
                result.AppendLine($"【{CellsHelper.CellIndexToName(row, col)}】");
                result.AppendLine($"Formula: {cell.Formula}");
                result.AppendLine($"Value: {cell.Value ?? "(calculating)"}");
                result.AppendLine();
            }
        }

        if (formulaCount == 0)
            result.AppendLine("No formulas found");
        else
            result.Insert(0, $"Total formulas: {formulaCount}\n\n");

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Gets the calculated result of a formula
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formula result value</returns>
    private async Task<string> GetFormulaResultAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell");
        var calculateBeforeRead = ArgumentHelper.GetBool(arguments, "calculateBeforeRead", true);

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (calculateBeforeRead) workbook.CalculateFormula();

        var result = $"Cell: {cell}\n";
        result += $"Formula: {cellObj.Formula ?? "(none)"}\n";

        var calculatedValue = cellObj.Value;

        if (!string.IsNullOrEmpty(cellObj.Formula))
            if (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str)))
            {
                calculatedValue = cellObj.DisplayStringValue;
                if (string.IsNullOrEmpty(calculatedValue?.ToString())) calculatedValue = cellObj.Formula;
            }

        result += $"Calculated Value: {calculatedValue ?? "(empty)"}\n";
        result += $"Value Type: {cellObj.Type}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Calculates all formulas in the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional outputPath</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based, not used but required for signature)</param>
    /// <returns>Success message</returns>
    private async Task<string> CalculateFormulasAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var cell = ArgumentHelper.GetStringNullable(arguments, "cell");

        using var workbook = new Workbook(path);

        if (!string.IsNullOrEmpty(cell))
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];

            if (!string.IsNullOrEmpty(cellObj.Formula))
            {
                // Calculate formula first
                workbook.CalculateFormula();
                // Access value to trigger calculation and ensure result is saved
                _ = cellObj.Value;
                // Convert formula result to value (replace formula with calculated value)
                cellObj.PutValue(cellObj.Value);
            }
        }
        else
        {
            workbook.CalculateFormula();
        }

        workbook.Save(outputPath);

        var result = "Formula calculation completed\n";
        result += $"Worksheet: {workbook.Worksheets[sheetIndex].Name}\n";
        if (!string.IsNullOrEmpty(cell)) result += $"Cell: {cell}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Sets an array formula to a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range and formula</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with range</returns>
    private async Task<string> SetArrayFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range");
        var formula = ArgumentHelper.GetString(arguments, "formula");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, range);

        // Remove curly braces if present (they're not needed)
        var cleanFormula = formula.TrimStart('{').TrimEnd('}');

        // Validate range dimensions
        if (rangeObj.RowCount <= 0 || rangeObj.ColumnCount <= 0)
            throw new ArgumentException(
                $"Invalid range dimensions: rows={rangeObj.RowCount}, columns={rangeObj.ColumnCount}");

        // Validate row and column indices
        if (rangeObj.FirstRow < 0 || rangeObj.FirstColumn < 0)
            throw new ArgumentException(
                $"Invalid range position: startRow={rangeObj.FirstRow}, startColumn={rangeObj.FirstColumn}");

        // The new API (FormulaParseOptions) is not available in this version of Aspose.Cells
        var firstCell = worksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

#pragma warning disable CS0618 // Type or member is obsolete
        // Set array formula using SetArrayFormula method
        // According to Aspose.Cells documentation, SetArrayFormula signature is:
        // SetArrayFormula(string arrayFormula, int rowNumber, int columnNumber)
        // where rowNumber and columnNumber are the number of rows and columns for the array

        // Formula should include '=' sign
        var formulaToSet = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;

        // Clear the range first
        for (var i = 0; i < rangeObj.RowCount; i++)
        for (var j = 0; j < rangeObj.ColumnCount; j++)
            worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        try
        {
            // Use SetArrayFormula with rowCount and columnCount (not startRow/startCol)
            // Signature: SetArrayFormula(formula, rowCount, columnCount)
            firstCell.SetArrayFormula(formulaToSet, rangeObj.RowCount, rangeObj.ColumnCount);

            // Calculate formulas to ensure array formula is processed
            workbook.CalculateFormula();

            // Check immediately if it's an array formula (optimization: avoid save/reload if already valid)
            if (firstCell.IsArrayFormula)
            {
                workbook.Save(outputPath);
                return await Task.FromResult($"Array formula set in range {range}: {outputPath}");
            }

            // Save and reload to verify (in case immediate check didn't detect it)
            workbook.Save(outputPath);
            using var verifyWorkbook = new Workbook(outputPath);
            var verifyWorksheet = verifyWorkbook.Worksheets[sheetIndex];
            var verifyCell = verifyWorksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

            if (verifyCell.IsArrayFormula)
                return await Task.FromResult($"Array formula set in range {range}: {outputPath}");

            // If SetArrayFormula with 2 parameters didn't work, try with 5 parameters
            throw new InvalidOperationException("SetArrayFormula with 2 parameters did not work");
        }
        catch (Exception ex)
        {
            // Try with 5 parameters: SetArrayFormula(formula, startRow, startCol, isR1C1, isLocal)
            try
            {
                // Reload for clean state
                using var retryWorkbook = new Workbook(path);
                var retryWorksheet = retryWorkbook.Worksheets[sheetIndex];
                var retryRangeObj = ExcelHelper.CreateRange(retryWorksheet.Cells, range);
                var retryFirstCell = retryWorksheet.Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];

                // Clear range
                for (var i = 0; i < retryRangeObj.RowCount; i++)
                for (var j = 0; j < retryRangeObj.ColumnCount; j++)
                    retryWorksheet.Cells[retryRangeObj.FirstRow + i, retryRangeObj.FirstColumn + j].PutValue("");

                // Try with 5 parameters
                var formulaWithoutEquals = cleanFormula.StartsWith("=") ? cleanFormula.Substring(1) : cleanFormula;
                retryFirstCell.SetArrayFormula(formulaWithoutEquals, retryRangeObj.FirstRow, retryRangeObj.FirstColumn,
                    false, false);

                retryWorkbook.CalculateFormula();
                retryWorkbook.Save(outputPath);

                // Verify
                using var verifyWorkbook = new Workbook(outputPath);
                var verifyWorksheet = verifyWorkbook.Worksheets[sheetIndex];
                var verifyCell = verifyWorksheet.Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];

                if (verifyCell.IsArrayFormula)
                    return await Task.FromResult($"Array formula set in range {range}: {outputPath}");

                throw new InvalidOperationException("SetArrayFormula with 5 parameters did not work");
            }
            catch (Exception ex2)
            {
                // If both methods fail, set regular formulas as fallback
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
                    return await Task.FromResult(
                        $"Formula set to range {range} (Note: This is not a true array formula): {outputPath}");
                }
                catch (Exception ex3)
                {
                    throw new ArgumentException(
                        $"Failed to set array formula. Range: {range}, Formula: {cleanFormula}.\nMethod 1 error: {ex.Message}\nMethod 2 error: {ex2.Message}\nMethod 3 error: {ex3.Message}",
                        ex);
                }
            }
        }
#pragma warning restore CS0618 // Type or member is obsolete
    }

    /// <summary>
    ///     Gets array formula information from a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Array formula information</returns>
    private async Task<string> GetArrayFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = ArgumentHelper.GetString(arguments, "cell");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        var sb = new StringBuilder();
        sb.AppendLine($"Cell: {cell}");

        // Use IsArrayFormula property to detect if cell contains an array formula
        if (cellObj.IsArrayFormula)
        {
            var formula = cellObj.Formula;
            sb.AppendLine($"Array Formula: {formula ?? "(empty)"}");

            // Try to get the array range
            try
            {
                // Find the array range by checking surrounding cells
                var startRow = cellObj.Row;
                var startCol = cellObj.Column;
                var endRow = startRow;
                var endCol = startCol;

                // Check cells to the right
                for (var col = startCol + 1; col < worksheet.Cells.MaxColumn + 1; col++)
                {
                    var testCell = worksheet.Cells[startRow, col];
                    if (testCell.IsArrayFormula && testCell.Formula == formula)
                        endCol = col;
                    else
                        break;
                }

                // Check cells below
                for (var row = startRow + 1; row < worksheet.Cells.MaxRow + 1; row++)
                {
                    var testCell = worksheet.Cells[row, startCol];
                    if (testCell.IsArrayFormula && testCell.Formula == formula)
                        endRow = row;
                    else
                        break;
                }

                var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                sb.AppendLine($"Array Range: {startCellName}:{endCellName}");
            }
            catch
            {
                sb.AppendLine("Array Range: Unable to determine");
            }
        }
        else
        {
            sb.AppendLine("No array formula found in this cell");
        }

        return await Task.FromResult(sb.ToString());
    }
}