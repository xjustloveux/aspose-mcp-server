using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel ranges (write, edit, get, clear, copy, move, copy_format)
///     Merges: ExcelWriteRangeTool, ExcelEditRangeTool, ExcelGetRangeTool, ExcelClearRangeTool,
///     ExcelCopyRangeTool, ExcelMoveRangeTool, ExcelCopyFormatTool
/// </summary>
[McpServerToolType]
public class ExcelRangeTool
{
    /// <summary>
    ///     Operation name for writing data to a range.
    /// </summary>
    private const string OperationWrite = "write";

    /// <summary>
    ///     Operation name for editing range data.
    /// </summary>
    private const string OperationEdit = "edit";

    /// <summary>
    ///     Operation name for getting range data.
    /// </summary>
    private const string OperationGet = "get";

    /// <summary>
    ///     Operation name for clearing range content.
    /// </summary>
    private const string OperationClear = "clear";

    /// <summary>
    ///     Operation name for copying a range.
    /// </summary>
    private const string OperationCopy = "copy";

    /// <summary>
    ///     Operation name for moving a range.
    /// </summary>
    private const string OperationMove = "move";

    /// <summary>
    ///     Operation name for copying format only.
    /// </summary>
    private const string OperationCopyFormat = "copy_format";

    /// <summary>
    ///     Text format number for Aspose.Cells style (formats cell as text).
    /// </summary>
    private const int TextFormatNumber = 49;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelRangeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelRangeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel range operation (write, edit, get, clear, copy, move, copy_format).
    /// </summary>
    /// <param name="operation">The operation to perform: write, edit, get, clear, copy, move, copy_format.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="sourceSheetIndex">Source sheet index (0-based, optional, for copy/move, default: same as sheetIndex).</param>
    /// <param name="destSheetIndex">Destination sheet index (0-based, optional, for copy/move, default: same as source).</param>
    /// <param name="startCell">Starting cell (e.g., 'A1', required for write).</param>
    /// <param name="range">
    ///     Source cell range (e.g., 'A1:C5', required for edit/get/clear operations, optional for
    ///     copy_format).
    /// </param>
    /// <param name="sourceRange">
    ///     Source range (e.g., 'A1:C5', required for copy/move, optional for copy_format as alternative
    ///     to range).
    /// </param>
    /// <param name="destCell">Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format).</param>
    /// <param name="destRange">Destination range (e.g., 'E1:G5', required for copy_format, or use destCell).</param>
    /// <param name="data">Data to write as JSON array.</param>
    /// <param name="clearRange">Clear range before writing (optional, for edit, default: false).</param>
    /// <param name="includeFormulas">Include formulas instead of values (optional, for get, default: false).</param>
    /// <param name="calculateFormulas">Recalculate all formulas before getting values (optional, for get, default: false).</param>
    /// <param name="includeFormat">Include format information (optional, for get, default: false).</param>
    /// <param name="clearContent">Clear cell content (optional, for clear, default: true).</param>
    /// <param name="clearFormat">Clear cell format (optional, for clear, default: false).</param>
    /// <param name="copyOptions">Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All').</param>
    /// <param name="copyValue">Copy cell values as well (optional, for copy_format, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_range")]
    [Description(@"Manage Excel ranges. Supports 7 operations: write, edit, get, clear, copy, move, copy_format.

Usage examples:
- Write range: excel_range(operation='write', path='book.xlsx', startCell='A1', data=[['A','B'],['C','D']])
- Edit range: excel_range(operation='edit', path='book.xlsx', range='A1:B2', data=[['X','Y']])
- Get range: excel_range(operation='get', path='book.xlsx', range='A1:B2')
- Clear range: excel_range(operation='clear', path='book.xlsx', range='A1:B2')
- Copy range: excel_range(operation='copy', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Move range: excel_range(operation='move', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Copy format: excel_range(operation='copy_format', path='book.xlsx', sourceRange='A1:B2', destCell='C1')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'write': Write data to range (required params: path, startCell, data)
- 'edit': Edit range data (required params: path, range, data)
- 'get': Get range data (required params: path, range)
- 'clear': Clear range (required params: path, range)
- 'copy': Copy range (required params: path, sourceRange, destCell)
- 'move': Move range (required params: path, sourceRange, destCell)
- 'copy_format': Copy format only (required params: path, range or sourceRange, destRange or destCell)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Source sheet index (0-based, optional, for copy/move, default: same as sheetIndex)")]
        int? sourceSheetIndex = null,
        [Description("Destination sheet index (0-based, optional, for copy/move, default: same as source)")]
        int? destSheetIndex = null,
        [Description("Starting cell (e.g., 'A1', required for write)")]
        string? startCell = null,
        [Description(
            "Source cell range (e.g., 'A1:C5', required for edit/get/clear operations, optional for copy_format)")]
        string? range = null,
        [Description(
            "Source range (e.g., 'A1:C5', required for copy/move, optional for copy_format as alternative to range)")]
        string? sourceRange = null,
        [Description("Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format)")]
        string? destCell = null,
        [Description("Destination range (e.g., 'E1:G5', required for copy_format, or use destCell)")]
        string? destRange = null,
        [Description(@"Data to write. Supports two formats:
1) 2D array: [['row1_col1', 'row1_col2'], ['row2_col1', 'row2_col2']]
2) Object array: [{'cell': 'A1', 'value': '10'}, {'cell': 'B1', 'value': '20'}]")]
        string? data = null,
        [Description("Clear range before writing (optional, for edit, default: false)")]
        bool clearRange = false,
        [Description("Include formulas instead of values (optional, for get, default: false)")]
        bool includeFormulas = false,
        [Description("Recalculate all formulas before getting values (optional, for get, default: false)")]
        bool calculateFormulas = false,
        [Description("Include format information (optional, for get, default: false)")]
        bool includeFormat = false,
        [Description("Clear cell content (optional, for clear, default: true)")]
        bool clearContent = true,
        [Description("Clear cell format (optional, for clear, default: false)")]
        bool clearFormat = false,
        [Description("Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All')")]
        string copyOptions = "All",
        [Description("Copy cell values as well (optional, for copy_format, default: false)")]
        bool copyValue = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLowerInvariant() switch
        {
            OperationWrite => WriteRange(ctx, outputPath, sheetIndex, startCell, data),
            OperationEdit => EditRange(ctx, outputPath, sheetIndex, range, data, clearRange),
            OperationGet => GetRange(ctx, sheetIndex, range, includeFormulas, calculateFormulas, includeFormat),
            OperationClear => ClearRange(ctx, outputPath, sheetIndex, range, clearContent, clearFormat),
            OperationCopy => CopyRange(ctx, outputPath, sheetIndex, sourceSheetIndex, destSheetIndex, sourceRange,
                destCell, copyOptions),
            OperationMove => MoveRange(ctx, outputPath, sheetIndex, sourceSheetIndex, destSheetIndex, sourceRange,
                destCell),
            OperationCopyFormat => CopyFormat(ctx, outputPath, sheetIndex, range ?? sourceRange, destRange ?? destCell,
                copyValue),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Writes data to a range starting at the specified cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="startCell">The starting cell address.</param>
    /// <param name="dataJson">The JSON data to write.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when startCell or dataJson is not provided, or data format is invalid.</exception>
    private static string WriteRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? startCell, string? dataJson)
    {
        if (string.IsNullOrEmpty(startCell))
            throw new ArgumentException("startCell is required for write operation");
        if (string.IsNullOrEmpty(dataJson))
            throw new ArgumentException("data is required for write operation");

        JsonArray dataArray;
        try
        {
            var parsed = JsonNode.Parse(dataJson);
            dataArray = parsed?.AsArray() ?? throw new ArgumentException("data must be a JSON array");
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON format for data: {ex.Message}");
        }

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var startCellObj = worksheet.Cells[startCell];
        var startRow = startCellObj.Row;
        var startCol = startCellObj.Column;

        var is2DArrayFormat = dataArray.All(item => item is JsonArray);

        if (is2DArrayFormat && dataArray.Count > 0)
        {
            var rowCount = dataArray.Count;
            var colCount = dataArray.Max(item => item is JsonArray arr ? arr.Count : 0);

            if (colCount > 0)
            {
                var data2D = new object[rowCount, colCount];

                for (var i = 0; i < rowCount; i++)
                    if (dataArray[i] is JsonArray rowArray)
                        for (var j = 0; j < colCount; j++)
                            if (j < rowArray.Count)
                                data2D[i, j] = ValueHelper.ParseValue(rowArray[j]?.GetValue<string>() ?? "");
                            else
                                data2D[i, j] = "";

                worksheet.Cells.ImportTwoDimensionArray(data2D, startRow, startCol);

                for (var i = 0; i < rowCount; i++)
                    if (dataArray[i] is JsonArray rowArray)
                        for (var j = 0; j < rowArray.Count; j++)
                        {
                            var cellValue = rowArray[j]?.GetValue<string>() ?? "";

                            var looksLikeCellRef = cellValue.Length >= 2 &&
                                                   char.IsLetter(cellValue[0]) &&
                                                   ((cellValue.Length is 2 && char.IsDigit(cellValue[1])) ||
                                                    (cellValue.Length is > 2 and <= 5 &&
                                                     cellValue.Skip(1).All(char.IsLetterOrDigit) &&
                                                     cellValue.Substring(1).Any(char.IsDigit) &&
                                                     !cellValue.Contains(' ') &&
                                                     !cellValue.Contains(':') &&
                                                     !cellValue.Contains('$')));

                            if (looksLikeCellRef && ValueHelper.ParseValue(cellValue) is string)
                            {
                                var cellObj = worksheet.Cells[startRow + i, startCol + j];
                                var style = workbook.CreateStyle();
                                style.Number = TextFormatNumber;
                                cellObj.SetStyle(style);
                                cellObj.PutValue(cellValue, true);
                            }
                        }
            }
        }
        else
        {
            for (var i = 0; i < dataArray.Count; i++)
            {
                var item = dataArray[i];

                if (item is JsonObject itemObj)
                {
                    var cellRef = itemObj["cell"]?.GetValue<string>();
                    var cellValue = itemObj["value"]?.GetValue<string>() ?? "";

                    if (!string.IsNullOrEmpty(cellRef))
                        ExcelHelper.SetCellValue(worksheet.Cells[cellRef], cellValue);
                }
                else
                {
                    throw new ArgumentException(
                        $"Invalid data format at index {i}. Expected array of arrays (2D) or array of objects with 'cell' and 'value' properties. Got: {item?.GetType().Name ?? "null"}");
                }
            }
        }

        ctx.Save(outputPath);
        return $"Data written to range starting at {startCell}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits data in an existing range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The range to edit.</param>
    /// <param name="dataJson">The JSON data to write.</param>
    /// <param name="clearRange">Whether to clear the range before writing.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range or dataJson is not provided, or data format is invalid.</exception>
    private static string EditRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, string? dataJson, bool clearRange)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for edit operation");
        if (string.IsNullOrEmpty(dataJson))
            throw new ArgumentException("data is required for edit operation");

        JsonArray dataArray;
        try
        {
            var parsed = JsonNode.Parse(dataJson);
            dataArray = parsed?.AsArray() ?? throw new ArgumentException("data must be a JSON array");
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid JSON format for data: {ex.Message}");
        }

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (clearRange)
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                cells[i, j].PutValue("");

        var startRow = cellRange.FirstRow;
        var startCol = cellRange.FirstColumn;

        for (var i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
                for (var j = 0; j < rowArray.Count; j++)
                {
                    var value = rowArray[j]?.GetValue<string>() ?? "";
                    ExcelHelper.SetCellValue(cells[startRow + i, startCol + j], value);
                }
        }

        ctx.Save(outputPath);
        return $"Range {range} edited. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets data from a range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The range to get data from.</param>
    /// <param name="includeFormulas">Whether to include formulas in the result.</param>
    /// <param name="calculateFormulas">Whether to calculate formulas before getting values.</param>
    /// <param name="includeFormat">Whether to include format information in the result.</param>
    /// <returns>A JSON string containing the range data.</returns>
    /// <exception cref="ArgumentException">Thrown when range is not provided.</exception>
    private static string GetRange(DocumentContext<Workbook> ctx, int sheetIndex, string? range,
        bool includeFormulas, bool calculateFormulas, bool includeFormat)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for get operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (calculateFormulas)
            workbook.CalculateFormula();

        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        List<object> cellList = [];
        for (var i = 0; i < cellRange.RowCount; i++)
        for (var j = 0; j < cellRange.ColumnCount; j++)
        {
            var cell = cells[cellRange.FirstRow + i, cellRange.FirstColumn + j];
            var cellRef = CellsHelper.CellIndexToName(cellRange.FirstRow + i, cellRange.FirstColumn + j);

            object? displayValue;
            string? formula = null;

            if (includeFormulas && !string.IsNullOrEmpty(cell.Formula)) formula = cell.Formula;

            if (!string.IsNullOrEmpty(cell.Formula))
            {
                displayValue = cell.Value;
                if (displayValue is CellValueType.IsError)
                    displayValue = cell.DisplayStringValue;
                if (displayValue == null || (displayValue is string str && string.IsNullOrEmpty(str)))
                {
                    displayValue = cell.DisplayStringValue;
                    if (string.IsNullOrEmpty(displayValue?.ToString())) displayValue = cell.Formula;
                }
            }
            else
            {
                displayValue = cell.Value ?? cell.DisplayStringValue;
            }

            if (includeFormat)
            {
                var style = cell.GetStyle();
                cellList.Add(new
                {
                    cell = cellRef,
                    value = displayValue?.ToString() ?? "(empty)",
                    formula,
                    format = new
                    {
                        fontName = style.Font.Name,
                        fontSize = style.Font.Size
                    }
                });
            }
            else
            {
                cellList.Add(new
                {
                    cell = cellRef,
                    value = displayValue?.ToString() ?? "(empty)",
                    formula
                });
            }
        }

        var result = new
        {
            range,
            rowCount = cellRange.RowCount,
            columnCount = cellRange.ColumnCount,
            count = cellList.Count,
            items = cellList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Clears content and/or format from a range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The range to clear.</param>
    /// <param name="clearContent">Whether to clear cell content.</param>
    /// <param name="clearFormat">Whether to clear cell format.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range is not provided.</exception>
    private static string ClearRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, bool clearContent, bool clearFormat)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for clear operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (clearContent && clearFormat)
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
            {
                cells[i, j].PutValue("");
                var defaultStyle = workbook.CreateStyle();
                cells[i, j].SetStyle(defaultStyle);
            }
        }
        else if (clearContent)
        {
            for (var i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            for (var j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                cells[i, j].PutValue("");
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellRange.ApplyStyle(defaultStyle, new StyleFlag { All = true });
        }

        ctx.Save(outputPath);
        return $"Range {range} cleared. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Copies a range to another location.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="sourceSheetIndex">The source worksheet index.</param>
    /// <param name="destSheetIndex">The destination worksheet index.</param>
    /// <param name="sourceRange">The source range to copy.</param>
    /// <param name="destCell">The destination cell.</param>
    /// <param name="copyOptions">The copy options (All, Values, Formats, Formulas).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sourceRange or destCell is not provided, or range format is invalid.</exception>
    private static string CopyRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? sourceSheetIndex, int? destSheetIndex, string? sourceRange, string? destCell, string copyOptions)
    {
        if (string.IsNullOrEmpty(sourceRange))
            throw new ArgumentException("sourceRange is required for copy operation");
        if (string.IsNullOrEmpty(destCell))
            throw new ArgumentException("destCell is required for copy operation");

        var workbook = ctx.Document;
        var srcSheetIdx = sourceSheetIndex ?? sheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = destSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        Range sourceRangeObj;
        Range destRangeObj;
        try
        {
            sourceRangeObj = sourceSheet.Cells.CreateRange(sourceRange);
            destRangeObj = destSheet.Cells.CreateRange(destCell);
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Invalid range format. Source range: '{sourceRange}', Destination cell: '{destCell}'. Range exceeds Excel limits (valid range: A1:XFD1048576). Error: {ex.Message}");
        }

        var copyOptionsEnum = copyOptions switch
        {
            "Values" => PasteType.Values,
            "Formats" => PasteType.Formats,
            "Formulas" => PasteType.Formulas,
            _ => PasteType.All
        };

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = copyOptionsEnum });

        ctx.Save(outputPath);
        return $"Range {sourceRange} copied to {destCell}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Moves a range to another location.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="sourceSheetIndex">The source worksheet index.</param>
    /// <param name="destSheetIndex">The destination worksheet index.</param>
    /// <param name="sourceRange">The source range to move.</param>
    /// <param name="destCell">The destination cell.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sourceRange or destCell is not provided.</exception>
    private static string MoveRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? sourceSheetIndex, int? destSheetIndex, string? sourceRange, string? destCell)
    {
        if (string.IsNullOrEmpty(sourceRange))
            throw new ArgumentException("sourceRange is required for move operation");
        if (string.IsNullOrEmpty(destCell))
            throw new ArgumentException("destCell is required for move operation");

        var workbook = ctx.Document;
        var srcSheetIdx = sourceSheetIndex ?? sheetIndex;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, srcSheetIdx);
        var destSheetIdx = destSheetIndex ?? srcSheetIdx;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        var sourceRangeObj = ExcelHelper.CreateRange(sourceSheet.Cells, sourceRange, "source range");
        var destRangeObj = ExcelHelper.CreateRange(destSheet.Cells, destCell, "destination cell");

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

        for (var i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
        for (var j = sourceRangeObj.FirstColumn;
             j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1;
             j++)
            sourceSheet.Cells[i, j].PutValue("");

        ctx.Save(outputPath);
        return $"Range {sourceRange} moved to {destCell}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Copies format (and optionally values) from source range to destination.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The source range to copy format from.</param>
    /// <param name="destTarget">The destination range or cell.</param>
    /// <param name="copyValue">Whether to also copy cell values.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range or destTarget is not provided.</exception>
    private static string CopyFormat(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, string? destTarget, bool copyValue)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range or sourceRange is required for copy_format operation");
        if (string.IsNullOrEmpty(destTarget))
            throw new ArgumentException(
                "Either destRange or destCell is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5' or destCell='E1'");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var sourceCellRange = ExcelHelper.CreateRange(cells, range, "source range");
        var destCellRange = ExcelHelper.CreateRange(cells, destTarget, "destination");

        var pasteOptions = new PasteOptions
        {
            PasteType = copyValue ? PasteType.All : PasteType.Formats
        };

        destCellRange.Copy(sourceCellRange, pasteOptions);

        ctx.Save(outputPath);

        var result = copyValue ? "Format with values copied" : "Format copied";
        result += $" from {range} to {destTarget}. {ctx.GetOutputMessage(outputPath)}";

        return result;
    }
}