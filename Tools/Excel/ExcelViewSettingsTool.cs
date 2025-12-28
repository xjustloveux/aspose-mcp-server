using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel view settings (zoom, gridlines, headers, zero values, etc.)
///     Merges: ExcelSetZoomTool, ExcelSetGridlinesVisibleTool, ExcelSetRowColumnHeadersVisibleTool,
///     ExcelSetZeroValuesVisibleTool, ExcelSetViewSettingsTool, ExcelSetColumnWidthTool, ExcelSetRowHeightTool,
///     ExcelSetSheetBackgroundTool, ExcelSetSheetTabColorTool
/// </summary>
public class ExcelViewSettingsTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage Excel view settings. Supports 14 operations: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, freeze_panes, split_window, auto_fit_column, auto_fit_row, show_formulas.

Usage examples:
- Set zoom: excel_view_settings(operation='set_zoom', path='book.xlsx', zoom=150)
- Set gridlines: excel_view_settings(operation='set_gridlines', path='book.xlsx', visible=false)
- Set column width: excel_view_settings(operation='set_column_width', path='book.xlsx', columnIndex=0, width=20)
- Set row height: excel_view_settings(operation='set_row_height', path='book.xlsx', rowIndex=0, height=30)
- Freeze panes: excel_view_settings(operation='freeze_panes', path='book.xlsx', freezeRow=1, freezeColumn=1)
- Split window: excel_view_settings(operation='split_window', path='book.xlsx', splitRow=5, splitColumn=2)
- Auto fit column: excel_view_settings(operation='auto_fit_column', path='book.xlsx', columnIndex=0)
- Show formulas: excel_view_settings(operation='show_formulas', path='book.xlsx', visible=true)
- Set all: excel_view_settings(operation='set_all', path='book.xlsx', zoom=150, showGridlines=true)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
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
- 'set_zoom': Set zoom level (required params: path, zoom)
- 'set_gridlines': Set gridlines visibility (required params: path, visible)
- 'set_headers': Set headers visibility (required params: path, visible)
- 'set_zero_values': Set zero values visibility (required params: path, visible)
- 'set_column_width': Set column width (required params: path, columnIndex, width)
- 'set_row_height': Set row height (required params: path, rowIndex, height)
- 'set_background': Set sheet background (required params: path, imagePath OR removeBackground)
- 'set_tab_color': Set tab color (required params: path, color)
- 'set_all': Set multiple settings (required params: path)
- 'freeze_panes': Freeze panes - fixed rows/columns stay visible when scrolling (required: freezeRow OR freezeColumn OR unfreeze)
- 'split_window': Split window - divide into independent scrollable panes (required: splitRow OR splitColumn OR removeSplit)
- 'auto_fit_column': Auto fit column width (required params: path, columnIndex; optional: startRow, endRow)
- 'auto_fit_row': Auto fit row height (required params: path, rowIndex; optional: startColumn, endColumn)
- 'show_formulas': Show/hide formulas (required params: path, visible)",
                @enum = new[]
                {
                    "set_zoom", "set_gridlines", "set_headers", "set_zero_values", "set_column_width", "set_row_height",
                    "set_background", "set_tab_color", "set_all", "freeze_panes", "split_window", "auto_fit_column",
                    "auto_fit_row", "show_formulas"
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
            zoom = new
            {
                type = "number",
                description = "Zoom percentage (10-400, required for set_zoom)"
            },
            visible = new
            {
                type = "boolean",
                description = "Visibility (required for set_gridlines/set_headers/set_zero_values)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based, required for set_column_width)"
            },
            width = new
            {
                type = "number",
                description = "Column width in characters (required for set_column_width)"
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based, required for set_row_height)"
            },
            height = new
            {
                type = "number",
                description = "Row height in points (required for set_row_height)"
            },
            imagePath = new
            {
                type = "string",
                description = "Background image file path (required for set_background)"
            },
            removeBackground = new
            {
                type = "boolean",
                description = "Remove background image (optional, for set_background, default: false)"
            },
            color = new
            {
                type = "string",
                description = "Color in hex format (e.g., 'FF0000' or '#FF0000', required for set_tab_color)"
            },
            showGridlines = new
            {
                type = "boolean",
                description = "Show gridlines (optional, for set_all)"
            },
            showRowColumnHeaders = new
            {
                type = "boolean",
                description = "Show row/column headers (optional, for set_all)"
            },
            showZeroValues = new
            {
                type = "boolean",
                description = "Show zero values (optional, for set_all)"
            },
            displayRightToLeft = new
            {
                type = "boolean",
                description = "Display right to left (optional, for set_all)"
            },
            freezeRow = new
            {
                type = "number",
                description = "Row index to freeze at (0-based, for freeze_panes, rows above this will be frozen)"
            },
            freezeColumn = new
            {
                type = "number",
                description =
                    "Column index to freeze at (0-based, for freeze_panes, columns left of this will be frozen)"
            },
            unfreeze = new
            {
                type = "boolean",
                description = "Remove frozen panes (optional, for freeze_panes, default: false)"
            },
            splitRow = new
            {
                type = "number",
                description = "Row position to split at in pixels (for split_window)"
            },
            splitColumn = new
            {
                type = "number",
                description = "Column position to split at in pixels (for split_window)"
            },
            removeSplit = new
            {
                type = "boolean",
                description = "Remove window split (optional, for split_window, default: false)"
            },
            startRow = new
            {
                type = "number",
                description = "Start row index for auto fit range (0-based, optional, for auto_fit_column)"
            },
            endRow = new
            {
                type = "number",
                description = "End row index for auto fit range (0-based, optional, for auto_fit_column)"
            },
            startColumn = new
            {
                type = "number",
                description = "Start column index for auto fit range (0-based, optional, for auto_fit_row)"
            },
            endColumn = new
            {
                type = "number",
                description = "End column index for auto fit range (0-based, optional, for auto_fit_row)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "set_zoom" => await SetZoomAsync(path, outputPath, sheetIndex, arguments),
            "set_gridlines" => await SetGridlinesAsync(path, outputPath, sheetIndex, arguments),
            "set_headers" => await SetHeadersAsync(path, outputPath, sheetIndex, arguments),
            "set_zero_values" => await SetZeroValuesAsync(path, outputPath, sheetIndex, arguments),
            "set_column_width" => await SetColumnWidthAsync(path, outputPath, sheetIndex, arguments),
            "set_row_height" => await SetRowHeightAsync(path, outputPath, sheetIndex, arguments),
            "set_background" => await SetBackgroundAsync(path, outputPath, sheetIndex, arguments),
            "set_tab_color" => await SetTabColorAsync(path, outputPath, sheetIndex, arguments),
            "set_all" => await SetAllAsync(path, outputPath, sheetIndex, arguments),
            "freeze_panes" => await FreezePanesAsync(path, outputPath, sheetIndex, arguments),
            "split_window" => await SplitWindowAsync(path, outputPath, sheetIndex, arguments),
            "auto_fit_column" => await AutoFitColumnAsync(path, outputPath, sheetIndex, arguments),
            "auto_fit_row" => await AutoFitRowAsync(path, outputPath, sheetIndex, arguments),
            "show_formulas" => await ShowFormulasAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets zoom level for the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing zoom (10-400)</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when zoom is not between 10 and 400</exception>
    private Task<string> SetZoomAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var zoom = ArgumentHelper.GetInt(arguments, "zoom");

            if (zoom < 10 || zoom > 400) throw new ArgumentException("Zoom must be between 10 and 400");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.Zoom = zoom;

            workbook.Save(outputPath);
            return $"Zoom level set to {zoom}% for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets gridlines visibility
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing isVisible</param>
    /// <returns>Success message</returns>
    private Task<string> SetGridlinesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var visible = ArgumentHelper.GetBool(arguments, "visible", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.IsGridlinesVisible = visible;

            workbook.Save(outputPath);
            return $"Gridlines visibility set to {(visible ? "visible" : "hidden")}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets row/column headers visibility
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing isVisible</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeadersAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var visible = ArgumentHelper.GetBool(arguments, "visible", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.IsRowColumnHeadersVisible = visible;

            workbook.Save(outputPath);
            return $"RowColumnHeaders visibility set to {(visible ? "visible" : "hidden")}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets zero values visibility
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing isVisible</param>
    /// <returns>Success message</returns>
    private Task<string> SetZeroValuesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var visible = ArgumentHelper.GetBool(arguments, "visible", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.DisplayZeros = visible;

            workbook.Save(outputPath);
            return $"Zero values visibility set to {(visible ? "visible" : "hidden")}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets column width
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing columnIndex, width</param>
    /// <returns>Success message</returns>
    private Task<string> SetColumnWidthAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var width = ArgumentHelper.GetDouble(arguments, "width");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            worksheet.Cells.SetColumnWidth(columnIndex, width);
            workbook.Save(outputPath);

            return $"Column {columnIndex} width set to {width} characters. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets row height
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing rowIndex, height</param>
    /// <returns>Success message</returns>
    private Task<string> SetRowHeightAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var height = ArgumentHelper.GetDouble(arguments, "height");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            worksheet.Cells.SetRowHeight(rowIndex, height);
            workbook.Save(outputPath);

            return $"Row {rowIndex} height set to {height} points. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets worksheet background image
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing imagePath or removeBackground</param>
    /// <returns>Success message</returns>
    /// <exception cref="FileNotFoundException">Thrown when image file is not found</exception>
    /// <exception cref="ArgumentException">Thrown when neither imagePath nor removeBackground is provided</exception>
    private Task<string> SetBackgroundAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");
            var removeBackground = ArgumentHelper.GetBool(arguments, "removeBackground", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (removeBackground)
            {
                worksheet.BackgroundImage = null;
            }
            else if (!string.IsNullOrEmpty(imagePath))
            {
                if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");
                var imageBytes = File.ReadAllBytes(imagePath);
                worksheet.BackgroundImage = imageBytes;
            }
            else
            {
                throw new ArgumentException("Either imagePath or removeBackground must be provided");
            }

            workbook.Save(outputPath);
            return removeBackground
                ? $"Background image removed from sheet {sheetIndex}. Output: {outputPath}"
                : $"Background image set for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets worksheet tab color
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing color</param>
    /// <returns>Success message</returns>
    private Task<string> SetTabColorAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var colorStr = ArgumentHelper.GetString(arguments, "color");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var color = ColorHelper.ParseColor(colorStr);

            worksheet.TabColor = color;

            workbook.Save(outputPath);
            return $"Sheet tab color set to {colorStr}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets all view settings at once
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing all view settings</param>
    /// <returns>Success message</returns>
    private Task<string> SetAllAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var zoom = ArgumentHelper.GetIntNullable(arguments, "zoom");
            var showGridlines = ArgumentHelper.GetBoolNullable(arguments, "showGridlines");
            var showRowColumnHeaders = ArgumentHelper.GetBoolNullable(arguments, "showRowColumnHeaders");
            var showZeroValues = ArgumentHelper.GetBoolNullable(arguments, "showZeroValues");
            var displayRightToLeft = ArgumentHelper.GetBoolNullable(arguments, "displayRightToLeft");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (zoom.HasValue)
            {
                if (zoom.Value < 10 || zoom.Value > 400) throw new ArgumentException("Zoom must be between 10 and 400");
                worksheet.Zoom = zoom.Value;
            }

            if (showGridlines.HasValue) worksheet.IsGridlinesVisible = showGridlines.Value;

            if (showRowColumnHeaders.HasValue) worksheet.IsRowColumnHeadersVisible = showRowColumnHeaders.Value;

            if (showZeroValues.HasValue) worksheet.DisplayZeros = showZeroValues.Value;

            if (displayRightToLeft.HasValue) worksheet.DisplayRightToLeft = displayRightToLeft.Value;

            workbook.Save(outputPath);
            return $"View settings updated for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Freezes panes at specified row and/or column
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing freezeRow, freezeColumn, or unfreeze</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when no freeze parameters are provided</exception>
    private Task<string> FreezePanesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var freezeRow = ArgumentHelper.GetIntNullable(arguments, "freezeRow");
            var freezeColumn = ArgumentHelper.GetIntNullable(arguments, "freezeColumn");
            var unfreeze = ArgumentHelper.GetBool(arguments, "unfreeze", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (unfreeze)
            {
                worksheet.UnFreezePanes();
            }
            else if (freezeRow.HasValue || freezeColumn.HasValue)
            {
                var row = freezeRow ?? 0;
                var col = freezeColumn ?? 0;
                worksheet.FreezePanes(row, col, row, col);
            }
            else
            {
                throw new ArgumentException("Either freezeRow, freezeColumn, or unfreeze must be provided");
            }

            workbook.Save(outputPath);
            return unfreeze
                ? $"Panes unfrozen for sheet {sheetIndex}. Output: {outputPath}"
                : $"Panes frozen at row {freezeRow ?? 0}, column {freezeColumn ?? 0} for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Splits worksheet window into independent scrollable panes
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing splitRow, splitColumn, or removeSplit</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when no split parameters are provided</exception>
    private Task<string> SplitWindowAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var splitRow = ArgumentHelper.GetIntNullable(arguments, "splitRow");
            var splitColumn = ArgumentHelper.GetIntNullable(arguments, "splitColumn");
            var removeSplit = ArgumentHelper.GetBool(arguments, "removeSplit", false);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (removeSplit)
            {
                worksheet.RemoveSplit();
            }
            else if (splitRow.HasValue || splitColumn.HasValue)
            {
                // Set active cell to determine split position
                var row = splitRow ?? 0;
                var col = splitColumn ?? 0;
                worksheet.ActiveCell = CellsHelper.CellIndexToName(row, col);
                worksheet.Split();
            }
            else
            {
                throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
            }

            workbook.Save(outputPath);
            return removeSplit
                ? $"Window split removed for sheet {sheetIndex}. Output: {outputPath}"
                : $"Window split at row {splitRow ?? 0}, column {splitColumn ?? 0} for sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Auto fits column width based on content
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing columnIndex and optional startRow, endRow</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when columnIndex is out of range</exception>
    private Task<string> AutoFitColumnAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var startRow = ArgumentHelper.GetIntNullable(arguments, "startRow");
            var endRow = ArgumentHelper.GetIntNullable(arguments, "endRow");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (startRow.HasValue && endRow.HasValue)
                worksheet.AutoFitColumn(columnIndex, startRow.Value, endRow.Value);
            else
                worksheet.AutoFitColumn(columnIndex);

            workbook.Save(outputPath);
            return $"Column {columnIndex} auto-fitted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Auto fits row height based on content
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing rowIndex and optional startColumn, endColumn</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when rowIndex is out of range</exception>
    private Task<string> AutoFitRowAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var startColumn = ArgumentHelper.GetIntNullable(arguments, "startColumn");
            var endColumn = ArgumentHelper.GetIntNullable(arguments, "endColumn");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (startColumn.HasValue && endColumn.HasValue)
                worksheet.AutoFitRow(rowIndex, startColumn.Value, endColumn.Value);
            else
                worksheet.AutoFitRow(rowIndex);

            workbook.Save(outputPath);
            return $"Row {rowIndex} auto-fitted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Shows or hides formulas in the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing visible</param>
    /// <returns>Success message</returns>
    private Task<string> ShowFormulasAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var visible = ArgumentHelper.GetBool(arguments, "visible", true);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            worksheet.ShowFormulas = visible;

            workbook.Save(outputPath);
            return $"Formulas {(visible ? "shown" : "hidden")} for sheet {sheetIndex}. Output: {outputPath}";
        });
    }
}