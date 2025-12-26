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
        @"Manage Excel view settings. Supports 10 operations: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, split_window.

Usage examples:
- Set zoom: excel_view_settings(operation='set_zoom', path='book.xlsx', zoom=150)
- Set gridlines: excel_view_settings(operation='set_gridlines', path='book.xlsx', visible=false)
- Set column width: excel_view_settings(operation='set_column_width', path='book.xlsx', columnIndex=0, width=20)
- Set row height: excel_view_settings(operation='set_row_height', path='book.xlsx', rowIndex=0, height=30)
- Set all: excel_view_settings(operation='set_all', path='book.xlsx', zoom=150, gridlinesVisible=true)";

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
- 'set_tab_color': Set tab color (required params: path, color; optional: sheetIndex, default: 0)
- 'set_all': Set multiple settings (required params: path)
- 'split_window': Split window (required params: path, splitRow OR splitColumn OR removeSplit)",
                @enum = new[]
                {
                    "set_zoom", "set_gridlines", "set_headers", "set_zero_values", "set_column_width", "set_row_height",
                    "set_background", "set_tab_color", "set_all", "split_window"
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
            splitRow = new
            {
                type = "number",
                description = "Row index to split at (0-based, optional, for split_window)"
            },
            splitColumn = new
            {
                type = "number",
                description = "Column index to split at (0-based, optional, for split_window)"
            },
            removeSplit = new
            {
                type = "boolean",
                description = "Remove split (optional, for split_window, default: false)"
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
            "split_window" => await SplitWindowAsync(path, outputPath, sheetIndex, arguments),
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
    ///     Sets worksheet background
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing imagePath or color</param>
    /// <returns>Success message</returns>
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
    ///     Splits worksheet window
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing splitRow, splitColumn</param>
    /// <returns>Success message</returns>
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
                // Split window - Use FreezePanes as alternative (requires 4 parameters)
                if (splitRow.HasValue && splitColumn.HasValue)
                    worksheet.FreezePanes(splitRow.Value + 1, splitColumn.Value + 1, splitRow.Value + 1,
                        splitColumn.Value + 1);
                else if (splitRow.HasValue)
                    worksheet.FreezePanes(splitRow.Value + 1, 0, splitRow.Value + 1, 0);
                else if (splitColumn.HasValue)
                    worksheet.FreezePanes(0, splitColumn.Value + 1, 0, splitColumn.Value + 1);
            }
            else
            {
                throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
            }

            workbook.Save(outputPath);
            return $"Window split {(removeSplit ? "removed" : "applied")} for sheet {sheetIndex}. Output: {outputPath}";
        });
    }
}