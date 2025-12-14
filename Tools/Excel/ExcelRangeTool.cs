using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel ranges (write, edit, get, clear, copy, move, copy_format)
/// Merges: ExcelWriteRangeTool, ExcelEditRangeTool, ExcelGetRangeTool, ExcelClearRangeTool, 
/// ExcelCopyRangeTool, ExcelMoveRangeTool, ExcelCopyFormatTool
/// </summary>
public class ExcelRangeTool : IAsposeTool
{
    public string Description => @"Manage Excel ranges. Supports 7 operations: write, edit, get, clear, copy, move, copy_format.

Usage examples:
- Write range: excel_range(operation='write', path='book.xlsx', startCell='A1', data=[['A','B'],['C','D']])
- Edit range: excel_range(operation='edit', path='book.xlsx', range='A1:B2', data=[['X','Y']])
- Get range: excel_range(operation='get', path='book.xlsx', range='A1:B2')
- Clear range: excel_range(operation='clear', path='book.xlsx', range='A1:B2')
- Copy range: excel_range(operation='copy', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Move range: excel_range(operation='move', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Copy format: excel_range(operation='copy_format', path='book.xlsx', sourceRange='A1:B2', destCell='C1')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'write': Write data to range (required params: path, startCell, data)
- 'edit': Edit range data (required params: path, range, data)
- 'get': Get range data (required params: path, range)
- 'clear': Clear range (required params: path, range)
- 'copy': Copy range (required params: path, sourceRange, destCell)
- 'move': Move range (required params: path, sourceRange, destCell)
- 'copy_format': Copy format only (required params: path, range, destRange or destCell)",
                @enum = new[] { "write", "edit", "get", "clear", "copy", "move", "copy_format" }
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
                description = "Source cell range (e.g., 'A1:C5', required for edit/get/clear/copy_format operations)"
            },
            sourceRange = new
            {
                type = "string",
                description = "Source range (e.g., 'A1:C5', required for copy/move)"
            },
            destCell = new
            {
                type = "string",
                description = "Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format as alternative to destRange)"
            },
            destRange = new
            {
                type = "string",
                description = "Destination range (e.g., 'E1:G5', required for copy_format, or use destCell for single cell)"
            },
            data = new
            {
                type = "array",
                description = "2D array of data to write",
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
            copyOptions = new
            {
                type = "string",
                description = "Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All')",
                @enum = new[] { "All", "Values", "Formats", "Formulas" }
            },
            copyValue = new
            {
                type = "boolean",
                description = "Copy cell values as well (optional, for copy_format, default: false)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "write" => await WriteRangeAsync(arguments, path, sheetIndex),
            "edit" => await EditRangeAsync(arguments, path, sheetIndex),
            "get" => await GetRangeAsync(arguments, path, sheetIndex),
            "clear" => await ClearRangeAsync(arguments, path, sheetIndex),
            "copy" => await CopyRangeAsync(arguments, path, sheetIndex),
            "move" => await MoveRangeAsync(arguments, path, sheetIndex),
            "copy_format" => await CopyFormatAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Writes data to a range starting at the specified cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing startCell and data array</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with range location</returns>
    private async Task<string> WriteRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var startCell = ArgumentHelper.GetString(arguments, "startCell", "startCell");
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required for write operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var startCellObj = worksheet.Cells[startCell];
        int startRow = startCellObj.Row;
        int startCol = startCellObj.Column;

        for (int i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
            {
                for (int j = 0; j < rowArray.Count; j++)
                {
                    var cellValue = rowArray[j]?.GetValue<string>() ?? "";
                    var cellObj = worksheet.Cells[startRow + i, startCol + j];
                    
                    // Parse value as number, boolean, or string
                    if (double.TryParse(cellValue, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double numValue))
                    {
                        cellObj.PutValue(numValue);
                    }
                    else if (bool.TryParse(cellValue, out bool boolValue))
                    {
                        cellObj.PutValue(boolValue);
                    }
                    else
                    {
                        cellObj.PutValue(cellValue);
                    }
                }
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Data written to range starting at {startCell}: {path}");
    }

    /// <summary>
    /// Edits data in an existing range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, data array, and optional clearRange</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with range information</returns>
    private async Task<string> EditRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var dataArray = arguments?["data"]?.AsArray() ?? throw new ArgumentException("data is required for edit operation");
        var clearRange = arguments?["clearRange"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (clearRange)
        {
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                }
            }
        }

        var startRow = cellRange.FirstRow;
        var startCol = cellRange.FirstColumn;

        for (int i = 0; i < dataArray.Count; i++)
        {
            var rowArray = dataArray[i]?.AsArray();
            if (rowArray != null)
            {
                for (int j = 0; j < rowArray.Count; j++)
                {
                    var value = rowArray[j]?.GetValue<string>() ?? "";
                    var cellObj = cells[startRow + i, startCol + j];
                    
                    // Parse value as number, boolean, or string
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
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {range} edited in sheet {sheetIndex}: {path}");
    }

    /// <summary>
    /// Gets data from a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range and optional includeFormulas, includeFormat</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with range data</returns>
    private async Task<string> GetRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var includeFormulas = arguments?["includeFormulas"]?.GetValue<bool?>() ?? false;
        var includeFormat = arguments?["includeFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        workbook.CalculateFormula();

        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        var sb = new StringBuilder();
        sb.AppendLine($"Range: {range}");
        sb.AppendLine($"Rows: {cellRange.RowCount}, Columns: {cellRange.ColumnCount}");
        sb.AppendLine();

        for (int i = 0; i < cellRange.RowCount; i++)
        {
            for (int j = 0; j < cellRange.ColumnCount; j++)
            {
                var cell = cells[cellRange.FirstRow + i, cellRange.FirstColumn + j];
                var cellRef = CellsHelper.CellIndexToName(cellRange.FirstRow + i, cellRange.FirstColumn + j);
                
                if (includeFormulas && !string.IsNullOrEmpty(cell.Formula))
                {
                    sb.Append($"{cellRef}: {cell.Formula}");
                }
                else
                {
                    object? displayValue;
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        displayValue = cell.Value;
                        if (displayValue is CellValueType cellType && cellType == CellValueType.IsError)
                        {
                            displayValue = cell.DisplayStringValue;
                        }
                        if (displayValue == null || (displayValue is string str && string.IsNullOrEmpty(str)))
                        {
                            displayValue = cell.DisplayStringValue;
                            if (string.IsNullOrEmpty(displayValue?.ToString()))
                            {
                                displayValue = cell.Formula;
                            }
                        }
                    }
                    else
                    {
                        displayValue = cell.Value ?? cell.DisplayStringValue;
                    }
                    
                    sb.Append($"{cellRef}: {displayValue ?? "(empty)"}");
                }

                if (includeFormat)
                {
                    var style = cell.GetStyle();
                    sb.Append($" [Font: {style.Font.Name}, Size: {style.Font.Size}]");
                }

                if (j < cellRange.ColumnCount - 1)
                {
                    sb.Append(" | ");
                }
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    /// Clears content and/or format from a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range and optional clearContent, clearFormat</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with cleared range</returns>
    private async Task<string> ClearRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var clearContent = arguments?["clearContent"]?.GetValue<bool?>() ?? true;
        var clearFormat = arguments?["clearFormat"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        if (clearContent && clearFormat)
        {
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                    var defaultStyle = workbook.CreateStyle();
                    cells[i, j].SetStyle(defaultStyle);
                }
            }
        }
        else if (clearContent)
        {
            for (int i = cellRange.FirstRow; i <= cellRange.FirstRow + cellRange.RowCount - 1; i++)
            {
                for (int j = cellRange.FirstColumn; j <= cellRange.FirstColumn + cellRange.ColumnCount - 1; j++)
                {
                    cells[i, j].PutValue("");
                }
            }
        }
        else if (clearFormat)
        {
            var defaultStyle = workbook.CreateStyle();
            cellRange.ApplyStyle(defaultStyle, new StyleFlag { All = true });
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {range} cleared in sheet {sheetIndex}: {path}");
    }

    /// <summary>
    /// Copies a range to another location
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourceRange, destCell, optional sourceSheetIndex, destSheetIndex, copyOptions</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with copy details</returns>
    private async Task<string> CopyRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int?>() ?? sheetIndex;
        var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange", "sourceRange");
        var destSheetIndex = arguments?["destSheetIndex"]?.GetValue<int?>();
        var destCell = ArgumentHelper.GetString(arguments, "destCell", "destCell");
        var copyOptions = arguments?["copyOptions"]?.GetValue<string>() ?? "All";

        using var workbook = new Workbook(path);
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
        var destSheetIdx = destSheetIndex ?? sourceSheetIndex;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        var sourceRangeObj = sourceSheet.Cells.CreateRange(sourceRange);
        var destRangeObj = destSheet.Cells.CreateRange(destCell);

        var copyOptionsEnum = copyOptions switch
        {
            "Values" => PasteType.Values,
            "Formats" => PasteType.Formats,
            "Formulas" => PasteType.Formulas,
            _ => PasteType.All
        };

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = copyOptionsEnum });

        workbook.Save(path);
        return await Task.FromResult($"Range {sourceRange} copied to {destCell} in sheet {destSheetIdx}: {path}");
    }

    /// <summary>
    /// Moves a range to another location
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourceRange, destCell, optional sourceSheetIndex, destSheetIndex</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with move details</returns>
    private async Task<string> MoveRangeAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int?>() ?? sheetIndex;
        var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange", "sourceRange");
        var destSheetIndex = arguments?["destSheetIndex"]?.GetValue<int?>();
        var destCell = ArgumentHelper.GetString(arguments, "destCell", "destCell");

        using var workbook = new Workbook(path);
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
        var destSheetIdx = destSheetIndex ?? sourceSheetIndex;
        var destSheet = ExcelHelper.GetWorksheet(workbook, destSheetIdx);

        var sourceRangeObj = sourceSheet.Cells.CreateRange(sourceRange);
        var destRangeObj = destSheet.Cells.CreateRange(destCell);

        destRangeObj.Copy(sourceRangeObj, new PasteOptions { PasteType = PasteType.All });

        for (int i = sourceRangeObj.FirstRow; i <= sourceRangeObj.FirstRow + sourceRangeObj.RowCount - 1; i++)
        {
            for (int j = sourceRangeObj.FirstColumn; j <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount - 1; j++)
            {
                sourceSheet.Cells[i, j].PutValue("");
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Range {sourceRange} moved to {destCell} in sheet {destSheetIdx}: {path}");
    }

    /// <summary>
    /// Copies format (and optionally values) from source range to destination
    /// </summary>
    /// <param name="arguments">JSON arguments containing range, destRange or destCell, optional copyValue</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with format copy details</returns>
    private async Task<string> CopyFormatAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();
        var destRange = arguments?["destRange"]?.GetValue<string>();
        var destCell = arguments?["destCell"]?.GetValue<string>();
        
        if (string.IsNullOrEmpty(range))
        {
            throw new ArgumentException("range (source range) is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5'");
        }
        
        if (string.IsNullOrEmpty(destRange) && string.IsNullOrEmpty(destCell))
        {
            throw new ArgumentException("Either destRange or destCell is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5' or destCell='E1'");
        }
        
        var copyValue = arguments?["copyValue"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        
        var sourceCellRange = cells.CreateRange(range);
        var destCellRange = cells.CreateRange(destRange ?? destCell!);

        var pasteOptions = new PasteOptions();
        pasteOptions.PasteType = copyValue ? PasteType.All : PasteType.Formats;
        pasteOptions.SkipBlanks = false;

        destCellRange.Copy(sourceCellRange, pasteOptions);

        workbook.Save(path);

        var result = "Format copied";
        if (copyValue)
        {
            result += " with values";
        }
        var destTarget = destRange ?? destCell!;
        result += $"\nSource range: {range}\nDestination: {destTarget}\nOutput: {path}";

        return await Task.FromResult(result);
    }
}

