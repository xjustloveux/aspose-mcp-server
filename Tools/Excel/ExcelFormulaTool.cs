using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel formulas (add, get, get_result, calculate, set_array, get_array).
///     Merges: ExcelAddFormulaTool, ExcelGetFormulaTool, ExcelGetFormulaResultTool,
///     ExcelCalculateFormulaTool, ExcelCalculateAllFormulasTool, ExcelSetArrayFormulaTool, ExcelGetArrayFormulaTool.
/// </summary>
[McpServerToolType]
public class ExcelFormulaTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelFormulaTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelFormulaTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_formula")]
    [Description(@"Manage Excel formulas. Supports 6 operations: add, get, get_result, calculate, set_array, get_array.

Usage examples:
- Add formula: excel_formula(operation='add', path='book.xlsx', cell='A1', formula='=SUM(B1:B10)')
- Get formula: excel_formula(operation='get', path='book.xlsx', cell='A1')
- Get result: excel_formula(operation='get_result', path='book.xlsx', cell='A1')
- Calculate: excel_formula(operation='calculate', path='book.xlsx')
- Set array formula: excel_formula(operation='set_array', path='book.xlsx', range='A1:A10', formula='=B1:B10*2')
- Get array formula: excel_formula(operation='get_array', path='book.xlsx', cell='A1')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a formula to a cell (required params: path, cell, formula)
- 'get': Get formula from a cell (required params: path, cell)
- 'get_result': Get formula result (required params: path, cell)
- 'calculate': Calculate all formulas (required params: path)
- 'set_array': Set array formula (required params: path, range, formula)
- 'get_array': Get array formula (required params: path, cell)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', required for add/get_result/get_array)")]
        string? cell = null,
        [Description("Cell range (e.g., 'A1:C10', optional for get, required for set_array)")]
        string? range = null,
        [Description("Formula (e.g., '=SUM(A1:A10)', required for add/set_array)")]
        string? formula = null,
        [Description("Calculate formulas before reading (optional, for get_result, default: true)")]
        bool calculateBeforeRead = true,
        [Description("Automatically calculate formulas after adding (optional, for add/set_array, default: true)")]
        bool autoCalculate = true)
    {
        return operation.ToLower() switch
        {
            "add" => AddFormula(path, sessionId, outputPath, sheetIndex, cell, formula, autoCalculate),
            "get" => GetFormulas(path, sessionId, sheetIndex, range),
            "get_result" => GetFormulaResult(path, sessionId, sheetIndex, cell, calculateBeforeRead),
            "calculate" => CalculateFormulas(path, sessionId, outputPath, sheetIndex),
            "set_array" => SetArrayFormula(path, sessionId, outputPath, sheetIndex, range, formula, autoCalculate),
            "get_array" => GetArrayFormula(path, sessionId, sheetIndex, cell),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a formula to a cell.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="formula">The formula to add.</param>
    /// <param name="autoCalculate">Whether to automatically calculate the formula.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when cell or formula is not provided.</exception>
    private string AddFormula(string? path, string? sessionId, string? outputPath, int sheetIndex, string? cell,
        string? formula, bool autoCalculate)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for add operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);

        var result = $"Formula added to {cell}: {formula}";
        if (!string.IsNullOrEmpty(warningMessage)) result += $".{warningMessage}";
        result += $". {ctx.GetOutputMessage(outputPath)}";
        return result;
    }

    /// <summary>
    ///     Gets all formulas from a range or the entire worksheet.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">Optional cell range to search for formulas.</param>
    /// <returns>A JSON string containing the formula information.</returns>
    /// <exception cref="ArgumentException">Thrown when the range format is invalid.</exception>
    private string GetFormulas(string? path, string? sessionId, int sheetIndex, string? range)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);
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

        List<object> formulaList = [];
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
    }

    /// <summary>
    ///     Gets the calculated result of a formula in a cell.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address containing the formula.</param>
    /// <param name="calculateBeforeRead">Whether to calculate formulas before reading the result.</param>
    /// <returns>A JSON string containing the formula result information.</returns>
    /// <exception cref="ArgumentException">Thrown when cell is not provided.</exception>
    private string GetFormulaResult(string? path, string? sessionId, int sheetIndex, string? cell,
        bool calculateBeforeRead)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for get_result operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var workbook = ctx.Document;
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
    }

    /// <summary>
    ///     Calculates all formulas in the workbook.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="_">Worksheet index (unused, calculates all sheets).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string CalculateFormulas(string? path, string? sessionId, string? outputPath, int _)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        ctx.Document.CalculateFormula();
        ctx.Save(outputPath);

        return $"Formulas calculated. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets an array formula for a range of cells.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range for the array formula.</param>
    /// <param name="formula">The array formula to set.</param>
    /// <param name="autoCalculate">Whether to automatically calculate the formula.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range or formula is not provided, or the range dimensions are invalid.</exception>
    private string SetArrayFormula(string? path, string? sessionId, string? outputPath, int sheetIndex, string? range,
        string? formula, bool autoCalculate)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for set_array operation");
        if (string.IsNullOrEmpty(formula))
            throw new ArgumentException("formula is required for set_array operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var workbook = ctx.Document;
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
                ctx.Save(outputPath);
                return $"Array formula set in range {range}. {ctx.GetOutputMessage(outputPath)}";
            }

            ctx.Save(outputPath);
            using var verifyWorkbook = new Workbook(outputPath ?? path!);
            var verifyCell = verifyWorkbook.Worksheets[sheetIndex].Cells[rangeObj.FirstRow, rangeObj.FirstColumn];

            if (verifyCell.IsArrayFormula)
                return $"Array formula set in range {range}. {ctx.GetOutputMessage(outputPath)}";

            throw new InvalidOperationException("SetArrayFormula with 2 parameters did not work");
        }
        catch (Exception ex)
        {
            try
            {
                using var retryWorkbook = new Workbook(path!);
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
                retryWorkbook.Save(outputPath ?? path!);

                using var verifyWorkbook = new Workbook(outputPath ?? path!);
                var verifyCell =
                    verifyWorkbook.Worksheets[sheetIndex].Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];

                if (verifyCell.IsArrayFormula)
                    return $"Array formula set in range {range}. Output: {outputPath ?? path}";

                throw new InvalidOperationException("SetArrayFormula with 5 parameters did not work");
            }
            catch (Exception ex2)
            {
                try
                {
                    using var fallbackWorkbook = new Workbook(path!);
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

                    fallbackWorkbook.Save(outputPath ?? path!);
                    return $"Formula set to range {range} (not a true array formula). Output: {outputPath ?? path}";
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
    }

    /// <summary>
    ///     Gets array formula information for a cell.
    /// </summary>
    /// <param name="path">The Excel file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address to check for array formula.</param>
    /// <returns>A JSON string containing the array formula information.</returns>
    /// <exception cref="ArgumentException">Thrown when cell is not provided.</exception>
    private string GetArrayFormula(string? path, string? sessionId, int sheetIndex, string? cell)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for get_array operation");

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);
        var worksheet = ExcelHelper.GetWorksheet(ctx.Document, sheetIndex);
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
            var rangeInfo = cellObj.GetArrayRange();
            var startCellName = CellsHelper.CellIndexToName(rangeInfo.StartRow, rangeInfo.StartColumn);
            var endCellName = CellsHelper.CellIndexToName(rangeInfo.EndRow, rangeInfo.EndColumn);
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
    }
}