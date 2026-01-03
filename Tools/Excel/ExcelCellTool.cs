using System.ComponentModel;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel cells (write, edit, get, clear)
/// </summary>
[McpServerToolType]
public class ExcelCellTool
{
    /// <summary>
    ///     Regex pattern for validating Excel cell addresses (e.g., A1, B2, AA100).
    /// </summary>
    private static readonly Regex CellAddressRegex = new(@"^[A-Za-z]{1,3}\d+$", RegexOptions.Compiled);

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelCellTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelCellTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_cell")]
    [Description(@"Manage Excel cells. Supports 4 operations: write, edit, get, clear.

Usage examples:
- Write cell: excel_cell(operation='write', path='book.xlsx', cell='A1', value='Hello')
- Edit cell: excel_cell(operation='edit', path='book.xlsx', cell='A1', value='Updated')
- Get cell: excel_cell(operation='get', path='book.xlsx', cell='A1')
- Clear cell: excel_cell(operation='clear', path='book.xlsx', cell='A1')")]
    public string Execute(
        [Description("Operation: write, edit, get, clear")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', 'B2', 'AA100')")]
        string? cell = null,
        [Description("Value to write")] string? value = null,
        [Description("Formula to set (optional, for edit, overrides value)")]
        string? formula = null,
        [Description("Clear cell value (optional, for edit)")]
        bool clearValue = false,
        [Description("Calculate formulas before reading value (optional, for get, default: false)")]
        bool calculateFormula = false,
        [Description("Include formula if present (optional, for get, default: true)")]
        bool includeFormula = true,
        [Description("Include format information (optional, for get, default: false)")]
        bool includeFormat = false,
        [Description("Clear cell content (optional, for clear, default: true)")]
        bool clearContent = true,
        [Description("Clear cell format (optional, for clear, default: false)")]
        bool clearFormat = false)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required");

        ValidateCellAddress(cell);

        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "write" => WriteCell(ctx, outputPath, sheetIndex, cell, value),
            "edit" => EditCell(ctx, outputPath, sheetIndex, cell, value, formula, clearValue),
            "get" => GetCell(ctx, sheetIndex, cell, calculateFormula, includeFormula, includeFormat),
            "clear" => ClearCell(ctx, outputPath, sheetIndex, cell, clearContent, clearFormat),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates the cell address format.
    /// </summary>
    /// <param name="cell">The cell address to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the cell address format is invalid.</exception>
    private static void ValidateCellAddress(string cell)
    {
        if (!CellAddressRegex.IsMatch(cell))
            throw new ArgumentException(
                $"Invalid cell address format: '{cell}'. Expected format like 'A1', 'B2', 'AA100'");
    }

    /// <summary>
    ///     Writes a value to the specified cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="value">The value to write.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string WriteCell(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string cell,
        string? value)
    {
        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("value is required for write operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        ExcelHelper.SetCellValue(cellObj, value);

        ctx.Save(outputPath);
        return $"Cell {cell} written with value '{value}' in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits the specified cell with a new value, formula, or clears it.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="value">The value to set.</param>
    /// <param name="formula">The formula to set.</param>
    /// <param name="clearValue">Whether to clear the cell value.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string EditCell(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string cell,
        string? value, string? formula, bool clearValue)
    {
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);
        return $"Cell {cell} edited in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets the value and properties of the specified cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="calculateFormula">Whether to calculate formulas before reading.</param>
    /// <param name="includeFormula">Whether to include formula in the result.</param>
    /// <param name="includeFormat">Whether to include format information in the result.</param>
    /// <returns>A JSON string containing the cell information.</returns>
    private static string GetCell(DocumentContext<Workbook> ctx, int sheetIndex, string cell, bool calculateFormula,
        bool includeFormula, bool includeFormat)
    {
        var workbook = ctx.Document;

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

    /// <summary>
    ///     Clears the content and/or format of the specified cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="clearContent">Whether to clear the cell content.</param>
    /// <param name="clearFormat">Whether to clear the cell format.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string ClearCell(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string cell,
        bool clearContent, bool clearFormat)
    {
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);
        return $"Cell {cell} cleared in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}