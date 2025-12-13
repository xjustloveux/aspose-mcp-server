using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel view settings (zoom, gridlines, headers, zero values, etc.)
/// Merges: ExcelSetZoomTool, ExcelSetGridlinesVisibleTool, ExcelSetRowColumnHeadersVisibleTool, 
/// ExcelSetZeroValuesVisibleTool, ExcelSetViewSettingsTool, ExcelSetColumnWidthTool, ExcelSetRowHeightTool,
/// ExcelSetSheetBackgroundTool, ExcelSetSheetTabColorTool
/// </summary>
public class ExcelViewSettingsTool : IAsposeTool
{
    public string Description => @"Manage Excel view settings. Supports 10 operations: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, split_window.

Usage examples:
- Set zoom: excel_view_settings(operation='set_zoom', path='book.xlsx', zoom=150)
- Set gridlines: excel_view_settings(operation='set_gridlines', path='book.xlsx', visible=false)
- Set column width: excel_view_settings(operation='set_column_width', path='book.xlsx', columnIndex=0, width=20)
- Set row height: excel_view_settings(operation='set_row_height', path='book.xlsx', rowIndex=0, height=30)
- Set all: excel_view_settings(operation='set_all', path='book.xlsx', zoom=150, gridlinesVisible=true)";

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
- 'set_background': Set sheet background (required params: path, imagePath)
- 'set_tab_color': Set tab color (required params: path, sheetIndex, color)
- 'set_all': Set multiple settings (required params: path)
- 'split_window': Split window (required params: path, rowIndex, columnIndex)",
                @enum = new[] { "set_zoom", "set_gridlines", "set_headers", "set_zero_values", "set_column_width", "set_row_height", "set_background", "set_tab_color", "set_all", "split_window" }
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "set_zoom" => await SetZoomAsync(arguments, path, sheetIndex),
            "set_gridlines" => await SetGridlinesAsync(arguments, path, sheetIndex),
            "set_headers" => await SetHeadersAsync(arguments, path, sheetIndex),
            "set_zero_values" => await SetZeroValuesAsync(arguments, path, sheetIndex),
            "set_column_width" => await SetColumnWidthAsync(arguments, path, sheetIndex),
            "set_row_height" => await SetRowHeightAsync(arguments, path, sheetIndex),
            "set_background" => await SetBackgroundAsync(arguments, path, sheetIndex),
            "set_tab_color" => await SetTabColorAsync(arguments, path, sheetIndex),
            "set_all" => await SetAllAsync(arguments, path, sheetIndex),
            "split_window" => await SplitWindowAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetZoomAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var zoom = arguments?["zoom"]?.GetValue<int>() ?? throw new ArgumentException("zoom is required for set_zoom operation");

        if (zoom < 10 || zoom > 400)
        {
            throw new ArgumentException("Zoom must be between 10 and 400");
        }

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Zoom = zoom;

        workbook.Save(path);
        return await Task.FromResult($"Zoom level set to {zoom}% for sheet {sheetIndex}: {path}");
    }

    private async Task<string> SetGridlinesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var visible = arguments?["visible"]?.GetValue<bool>() ?? throw new ArgumentException("visible is required for set_gridlines operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.IsGridlinesVisible = visible;

        workbook.Save(path);
        return await Task.FromResult($"Gridlines visibility set to {(visible ? "visible" : "hidden")}: {path}");
    }

    private async Task<string> SetHeadersAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var visible = arguments?["visible"]?.GetValue<bool>() ?? throw new ArgumentException("visible is required for set_headers operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.IsRowColumnHeadersVisible = visible;

        workbook.Save(path);
        return await Task.FromResult($"RowColumnHeaders visibility set to {(visible ? "visible" : "hidden")}: {path}");
    }

    private async Task<string> SetZeroValuesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var visible = arguments?["visible"]?.GetValue<bool>() ?? throw new ArgumentException("visible is required for set_zero_values operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.DisplayZeros = visible;

        workbook.Save(path);
        return await Task.FromResult($"Zero values visibility set to {(visible ? "visible" : "hidden")}: {path}");
    }

    private async Task<string> SetColumnWidthAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required for set_column_width operation");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required for set_column_width operation");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.Cells.SetColumnWidth(columnIndex, width);
        workbook.Save(path);

        return await Task.FromResult($"第 {columnIndex} 列寬度已設定為 {width} 字元: {path}");
    }

    private async Task<string> SetRowHeightAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required for set_row_height operation");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required for set_row_height operation");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.Cells.SetRowHeight(rowIndex, height);
        workbook.Save(path);

        return await Task.FromResult($"第 {rowIndex} 行高度已設定為 {height} 點: {path}");
    }

    private async Task<string> SetBackgroundAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var imagePath = arguments?["imagePath"]?.GetValue<string>();
        var removeBackground = arguments?["removeBackground"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (removeBackground)
        {
            worksheet.BackgroundImage = null;
        }
        else if (!string.IsNullOrEmpty(imagePath))
        {
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }
            var imageBytes = File.ReadAllBytes(imagePath);
            worksheet.BackgroundImage = imageBytes;
        }
        else
        {
            throw new ArgumentException("Either imagePath or removeBackground must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(removeBackground 
            ? $"Background image removed from sheet {sheetIndex}: {path}"
            : $"Background image set for sheet {sheetIndex}: {path}");
    }

    private async Task<string> SetTabColorAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var colorStr = arguments?["color"]?.GetValue<string>() ?? throw new ArgumentException("color is required for set_tab_color operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        System.Drawing.Color color;
        if (colorStr.StartsWith("#"))
        {
            colorStr = colorStr.Substring(1);
        }

        if (colorStr.Length == 6)
        {
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            color = System.Drawing.Color.FromArgb(r, g, b);
        }
        else
        {
            color = System.Drawing.Color.FromName(colorStr);
        }

        worksheet.TabColor = color;

        workbook.Save(path);
        return await Task.FromResult($"Sheet tab color set to {colorStr}: {path}");
    }

    private async Task<string> SetAllAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var zoom = arguments?["zoom"]?.GetValue<int?>();
        var showGridlines = arguments?["showGridlines"]?.GetValue<bool?>();
        var showRowColumnHeaders = arguments?["showRowColumnHeaders"]?.GetValue<bool?>();
        var showZeroValues = arguments?["showZeroValues"]?.GetValue<bool?>();
        var displayRightToLeft = arguments?["displayRightToLeft"]?.GetValue<bool?>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (zoom.HasValue)
        {
            if (zoom.Value < 10 || zoom.Value > 400)
            {
                throw new ArgumentException("Zoom must be between 10 and 400");
            }
            worksheet.Zoom = zoom.Value;
        }

        if (showGridlines.HasValue)
        {
            worksheet.IsGridlinesVisible = showGridlines.Value;
        }

        if (showRowColumnHeaders.HasValue)
        {
            worksheet.IsRowColumnHeadersVisible = showRowColumnHeaders.Value;
        }

        if (showZeroValues.HasValue)
        {
            worksheet.DisplayZeros = showZeroValues.Value;
        }

        if (displayRightToLeft.HasValue)
        {
            worksheet.DisplayRightToLeft = displayRightToLeft.Value;
        }

        workbook.Save(path);
        return await Task.FromResult($"View settings updated for sheet {sheetIndex}: {path}");
    }

    private async Task<string> SplitWindowAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var splitRow = arguments?["splitRow"]?.GetValue<int?>();
        var splitColumn = arguments?["splitColumn"]?.GetValue<int?>();
        var removeSplit = arguments?["removeSplit"]?.GetValue<bool?>() ?? false;

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
            {
                worksheet.FreezePanes(splitRow.Value + 1, splitColumn.Value + 1, splitRow.Value + 1, splitColumn.Value + 1);
            }
            else if (splitRow.HasValue)
            {
                worksheet.FreezePanes(splitRow.Value + 1, 0, splitRow.Value + 1, 0);
            }
            else if (splitColumn.HasValue)
            {
                worksheet.FreezePanes(0, splitColumn.Value + 1, 0, splitColumn.Value + 1);
            }
        }
        else
        {
            throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult($"Window split {(removeSplit ? "removed" : "applied")} for sheet {sheetIndex}: {path}");
    }
}

