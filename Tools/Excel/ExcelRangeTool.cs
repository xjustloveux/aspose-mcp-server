using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel ranges (write, edit, get, clear, copy, move, copy_format)
///     Merges: ExcelWriteRangeTool, ExcelEditRangeTool, ExcelGetRangeTool, ExcelClearRangeTool,
///     ExcelCopyRangeTool, ExcelMoveRangeTool, ExcelCopyFormatTool
/// </summary>
public class ExcelRangeTool : IAsposeTool
{
    private const string OperationWrite = "write";
    private const string OperationEdit = "edit";
    private const string OperationGet = "get";
    private const string OperationClear = "clear";
    private const string OperationCopy = "copy";
    private const string OperationMove = "move";
    private const string OperationCopyFormat = "copy_format";

    /// <summary>
    ///     Text format number for Aspose.Cells style (formats cell as text).
    /// </summary>
    private const int TextFormatNumber = 49;

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel ranges. Supports 7 operations: write, edit, get, clear, copy, move, copy_format.

Usage examples:
- Write range: excel_range(operation='write', path='book.xlsx', startCell='A1', data=[['A','B'],['C','D']])
- Edit range: excel_range(operation='edit', path='book.xlsx', range='A1:B2', data=[['X','Y']])
- Get range: excel_range(operation='get', path='book.xlsx', range='A1:B2')
- Clear range: excel_range(operation='clear', path='book.xlsx', range='A1:B2')
- Copy range: excel_range(operation='copy', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Move range: excel_range(operation='move', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Copy format: excel_range(operation='copy_format', path='book.xlsx', sourceRange='A1:B2', destCell='C1')";

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
                description = $@"Operation to perform.
- '{OperationWrite}': Write data to range (required params: path, startCell, data)
- '{OperationEdit}': Edit range data (required params: path, range, data)
- '{OperationGet}': Get range data (required params: path, range)
- '{OperationClear}': Clear range (required params: path, range)
- '{OperationCopy}': Copy range (required params: path, sourceRange, destCell)
- '{OperationMove}': Move range (required params: path, sourceRange, destCell)
- '{OperationCopyFormat}': Copy format only (required params: path, range or sourceRange, destRange or destCell)",
                @enum = new[]
                {
                    OperationWrite, OperationEdit, OperationGet, OperationClear, OperationCopy, OperationMove,
                    OperationCopyFormat
                }
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
            sourceSheetIndex = new
            {
                type = "number",
                description = "Source sheet index (0-based, optional, for copy/move, default: same as sheetIndex)"
            },
            destSheetIndex = new
            {
                type = "number",
                description = "Destination sheet index (0-based, optional, for copy/move, default: same as source)"
            },
            startCell = new
            {
                type = "string",
                description = "Starting cell (e.g., 'A1', required for write)"
            },
            range = new
            {
                type = "string",
                description =
                    "Source cell range (e.g., 'A1:C5', required for edit/get/clear operations, optional for copy_format - can use sourceRange instead)"
            },
            sourceRange = new
            {
                type = "string",
                description =
                    "Source range (e.g., 'A1:C5', required for copy/move, optional for copy_format as alternative to range)"
            },
            destCell = new
            {
                type = "string",
                description =
                    "Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format as alternative to destRange)"
            },
            destRange = new
            {
                type = "string",
                description =
                    "Destination range (e.g., 'E1:G5', required for copy_format, or use destCell for single cell)"
            },
            data = new
            {
                type = "array",
                description =
                    @"Data to write. Supports two formats:
1) 2D array format:
   - Structure: [[row1_col1, row1_col2, row1_col3, ...], [row2_col1, row2_col2, ...], ...]
   - Each sub-array represents one row, starting from startCell
   - Example: [[""A"", ""B"", ""C""], [""1"", ""2"", ""3""]] writes:
     * Row 1: A, B, C
     * Row 2: 1, 2, 3
   - For text values that look like cell references (e.g., ""A2""), use single quote prefix: ""'A2""
   - Example with single quote: [[""'A2"", ""10"", ""20"", ""30""]] writes one row: A2 (text), 10, 20, 30
2) Object array format:
   - Structure: [{""cell"": ""A1"", ""value"": ""10""}, {""cell"": ""B1"", ""value"": ""20""}]
   - Each object specifies exact cell location and value
   - startCell parameter is not needed when using object array format",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            clearRange = new
            {
                type = "boolean",
                description = "Clear range before writing (optional, for edit, default: false)"
            },
            includeFormulas = new
            {
                type = "boolean",
                description = "Include formulas instead of values (optional, for get, default: false)"
            },
            calculateFormulas = new
            {
                type = "boolean",
                description =
                    "Recalculate all formulas before getting values (optional, for get, default: false). Set to true if you need up-to-date formula results."
            },
            includeFormat = new
            {
                type = "boolean",
                description = "Include format information (optional, for get, default: false)"
            },
            clearContent = new
            {
                type = "boolean",
                description =
                    "Clear cell content (required for clear operation, default: true). Set clearContent=true to clear cell content, or clearContent=false to keep content."
            },
            clearFormat = new
            {
                type = "boolean",
                description = "Clear cell format (optional, for clear, default: false)"
            },
            copyOptions = new
            {
                type = "string",
                description =
                    "Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All')",
                @enum = new[] { "All", "Values", "Formats", "Formulas" }
            },
            copyValue = new
            {
                type = "boolean",
                description = "Copy cell values as well (optional, for copy_format, default: false)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for write/edit/clear/copy/move/copy_format operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLowerInvariant() switch
        {
            OperationWrite => await WriteRangeAsync(path, outputPath, sheetIndex, arguments),
            OperationEdit => await EditRangeAsync(path, outputPath, sheetIndex, arguments),
            OperationGet => await GetRangeAsync(path, sheetIndex, arguments),
            OperationClear => await ClearRangeAsync(path, outputPath, sheetIndex, arguments),
            OperationCopy => await CopyRangeAsync(path, outputPath, sheetIndex, arguments),
            OperationMove => await MoveRangeAsync(path, outputPath, sheetIndex, arguments),
            OperationCopyFormat => await CopyFormatAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Writes data to a range starting at the specified cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing startCell and data array.</param>
    /// <returns>Success message with range location.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or data format is invalid.</exception>
    private Task<string> WriteRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startCell = ArgumentHelper.GetString(arguments, "startCell");
            var dataArray = ArgumentHelper.GetArray(arguments, "data");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var startCellObj = worksheet.Cells[startCell];
            var startRow = startCellObj.Row;
            var startCol = startCellObj.Column;

            // Check if all items are arrays (2D array format) to use ImportTwoDimensionArray
            var is2DArrayFormat = dataArray.All(item => item is JsonArray);

            if (is2DArrayFormat && dataArray.Count > 0)
            {
                // 2D array format: use ImportTwoDimensionArray to avoid cell reference interpretation issues
                var rowCount = dataArray.Count;
                var colCount = dataArray.Max(item => item is JsonArray arr ? arr.Count : 0);

                if (colCount > 0)
                {
                    var data2D = new object[rowCount, colCount];

                    for (var i = 0; i < rowCount; i++)
                        if (dataArray[i] is JsonArray rowArray)
                            for (var j = 0; j < colCount; j++)
                                if (j < rowArray.Count)
                                    data2D[i, j] = ArgumentHelper.ParseValue(rowArray[j]?.GetValue<string>() ?? "");
                                else
                                    data2D[i, j] = "";

                    worksheet.Cells.ImportTwoDimensionArray(data2D, startRow, startCol);

                    // Set text format for values that look like cell references
                    for (var i = 0; i < rowCount; i++)
                        if (dataArray[i] is JsonArray rowArray)
                            for (var j = 0; j < rowArray.Count; j++)
                            {
                                var cellValue = rowArray[j]?.GetValue<string>() ?? "";

                                // Check if value looks like a cell reference (e.g., "A2", "B10")
                                var looksLikeCellRef = cellValue.Length >= 2 &&
                                                       char.IsLetter(cellValue[0]) &&
                                                       ((cellValue.Length is 2 && char.IsDigit(cellValue[1])) ||
                                                        (cellValue.Length is > 2 and <= 5 &&
                                                         cellValue.Skip(1).All(char.IsLetterOrDigit) &&
                                                         cellValue.Substring(1).Any(char.IsDigit) &&
                                                         !cellValue.Contains(" ") &&
                                                         !cellValue.Contains(":") &&
                                                         !cellValue.Contains("$")));

                                // If it looks like a cell ref and wasn't parsed as number/bool/date, force text format
                                if (looksLikeCellRef && ArgumentHelper.ParseValue(cellValue) is string)
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
                // Object format: [{"cell": "A1", "value": "10"}, {"cell": "B1", "value": "20"}]
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

            workbook.Save(outputPath);
            return $"Data written to range starting at {startCell}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits data in an existing range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range, data array, and optional clearRange.</param>
    /// <returns>Success message with range information.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range format is invalid.</exception>
    private Task<string> EditRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var dataArray = ArgumentHelper.GetArray(arguments, "data");
            var clearRange = ArgumentHelper.GetBool(arguments, "clearRange", false);

            using var workbook = new Workbook(path);
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

            workbook.Save(outputPath);
            return $"Range {range} edited. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets data from a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range and optional includeFormulas, includeFormat, calculateFormulas.</param>
    /// <returns>JSON string with range data.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range format is invalid.</exception>
    private Task<string> GetRangeAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var includeFormulas = ArgumentHelper.GetBool(arguments, "includeFormulas", false);
            var includeFormat = ArgumentHelper.GetBool(arguments, "includeFormat", false);
            var calculateFormulas = ArgumentHelper.GetBool(arguments, "calculateFormulas", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (calculateFormulas)
                workbook.CalculateFormula();

            var cells = worksheet.Cells;

            var cellRange = ExcelHelper.CreateRange(cells, range);

            var cellList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Clears content and/or format from a range.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range and optional clearContent, clearFormat.</param>
    /// <returns>Success message with cleared range.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range format is invalid.</exception>
    private Task<string> ClearRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetString(arguments, "range");
            var clearContent = ArgumentHelper.GetBool(arguments, "clearContent", true);
            var clearFormat = ArgumentHelper.GetBool(arguments, "clearFormat", false);

            using var workbook = new Workbook(path);
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

            workbook.Save(outputPath);
            return $"Range {range} cleared. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies a range to another location.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">
    ///     JSON arguments containing sourceRange, destCell, optional sourceSheetIndex, destSheetIndex,
    ///     copyOptions.
    /// </param>
    /// <returns>Success message with copy details.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range format is invalid.</exception>
    private Task<string> CopyRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sourceSheetIndex = ArgumentHelper.GetInt(arguments, "sourceSheetIndex", sheetIndex);
            var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange");
            var destSheetIndex = ArgumentHelper.GetIntNullable(arguments, "destSheetIndex");
            var destCell = ArgumentHelper.GetString(arguments, "destCell");
            var copyOptions = ArgumentHelper.GetString(arguments, "copyOptions", "All");

            using var workbook = new Workbook(path);
            var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
            var destSheetIdx = destSheetIndex ?? sourceSheetIndex;
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

            workbook.Save(outputPath);
            return $"Range {sourceRange} copied to {destCell}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Moves a range to another location.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing sourceRange, destCell, optional sourceSheetIndex, destSheetIndex.</param>
    /// <returns>Success message with move details.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range or range format is invalid.</exception>
    private Task<string> MoveRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sourceSheetIndex = ArgumentHelper.GetInt(arguments, "sourceSheetIndex", sheetIndex);
            var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange");
            var destSheetIndex = ArgumentHelper.GetIntNullable(arguments, "destSheetIndex");
            var destCell = ArgumentHelper.GetString(arguments, "destCell");

            using var workbook = new Workbook(path);
            var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
            var destSheetIdx = destSheetIndex ?? sourceSheetIndex;
            var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

            var sourceRangeObj = ExcelHelper.CreateRange(sourceSheet.Cells, sourceRange, "source range");
            var destRangeObj = ExcelHelper.CreateRange(destSheet.Cells, destCell, "destination cell");

            destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

            for (var i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
            for (var j = sourceRangeObj.FirstColumn;
                 j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1;
                 j++)
                sourceSheet.Cells[i, j].PutValue("");

            workbook.Save(outputPath);
            return $"Range {sourceRange} moved to {destCell}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies format (and optionally values) from source range to destination.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing range or sourceRange, destRange or destCell, optional copyValue.</param>
    /// <returns>Success message with format copy details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when sheetIndex is out of range, range format is invalid, or
    ///     destRange/destCell is missing.
    /// </exception>
    private Task<string> CopyFormatAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Support both 'range' and 'sourceRange' for consistency with copy/move operations
            var range = ArgumentHelper.GetString(arguments, "range", "sourceRange", "range or sourceRange");
            var destRange = ArgumentHelper.GetString(arguments, "destRange", false);
            var destCell = ArgumentHelper.GetString(arguments, "destCell", false);

            if (string.IsNullOrEmpty(destRange) && string.IsNullOrEmpty(destCell))
                throw new ArgumentException(
                    "Either destRange or destCell is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5' or destCell='E1'");

            var copyValue = ArgumentHelper.GetBool(arguments, "copyValue", false);

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            var destTarget = destRange ?? destCell!;
            var sourceCellRange = ExcelHelper.CreateRange(cells, range, "source range");
            var destCellRange = ExcelHelper.CreateRange(cells, destTarget, "destination");

            var pasteOptions = new PasteOptions
            {
                PasteType = copyValue ? PasteType.All : PasteType.Formats
            };

            destCellRange.Copy(sourceCellRange, pasteOptions);

            workbook.Save(outputPath);

            var result = copyValue ? "Format with values copied" : "Format copied";
            result += $" from {range} to {destTarget}. Output: {outputPath}";

            return result;
        });
    }
}