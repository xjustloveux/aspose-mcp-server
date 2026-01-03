using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel view settings (zoom, gridlines, headers, zero values, etc.)
///     Merges: ExcelSetZoomTool, ExcelSetGridlinesVisibleTool, ExcelSetRowColumnHeadersVisibleTool,
///     ExcelSetZeroValuesVisibleTool, ExcelSetViewSettingsTool, ExcelSetColumnWidthTool, ExcelSetRowHeightTool,
///     ExcelSetSheetBackgroundTool, ExcelSetSheetTabColorTool
/// </summary>
[McpServerToolType]
public class ExcelViewSettingsTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelViewSettingsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelViewSettingsTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_view_settings")]
    [Description(
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
- Set all: excel_view_settings(operation='set_all', path='book.xlsx', zoom=150, showGridlines=true)")]
    public string Execute(
        [Description(
            "Operation: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, freeze_panes, split_window, auto_fit_column, auto_fit_row, show_formulas")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Zoom percentage (10-400, required for set_zoom)")]
        int zoom = 100,
        [Description("Visibility (required for set_gridlines/set_headers/set_zero_values/show_formulas)")]
        bool visible = true,
        [Description("Column index (0-based, required for set_column_width/auto_fit_column)")]
        int columnIndex = 0,
        [Description("Column width in characters (required for set_column_width)")]
        double width = 8.43,
        [Description("Row index (0-based, required for set_row_height/auto_fit_row)")]
        int rowIndex = 0,
        [Description("Row height in points (required for set_row_height)")]
        double height = 15,
        [Description("Background image file path (required for set_background)")]
        string? imagePath = null,
        [Description("Remove background image (for set_background)")]
        bool removeBackground = false,
        [Description("Color in hex format (e.g., '#FF0000', required for set_tab_color)")]
        string? color = null,
        [Description("Show gridlines (for set_all)")]
        bool? showGridlines = null,
        [Description("Show row/column headers (for set_all)")]
        bool? showRowColumnHeaders = null,
        [Description("Show zero values (for set_all)")]
        bool? showZeroValues = null,
        [Description("Display right to left (for set_all)")]
        bool? displayRightToLeft = null,
        [Description("Row index to freeze at (0-based, for freeze_panes)")]
        int? freezeRow = null,
        [Description("Column index to freeze at (0-based, for freeze_panes)")]
        int? freezeColumn = null,
        [Description("Remove frozen panes (for freeze_panes)")]
        bool unfreeze = false,
        [Description("Row position to split at in pixels (for split_window)")]
        int? splitRow = null,
        [Description("Column position to split at in pixels (for split_window)")]
        int? splitColumn = null,
        [Description("Remove window split (for split_window)")]
        bool removeSplit = false,
        [Description("Start row index for auto fit range (0-based, for auto_fit_column)")]
        int? startRow = null,
        [Description("End row index for auto fit range (0-based, for auto_fit_column)")]
        int? endRow = null,
        [Description("Start column index for auto fit range (0-based, for auto_fit_row)")]
        int? startColumn = null,
        [Description("End column index for auto fit range (0-based, for auto_fit_row)")]
        int? endColumn = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set_zoom" => SetZoom(ctx, outputPath, sheetIndex, zoom),
            "set_gridlines" => SetGridlines(ctx, outputPath, sheetIndex, visible),
            "set_headers" => SetHeaders(ctx, outputPath, sheetIndex, visible),
            "set_zero_values" => SetZeroValues(ctx, outputPath, sheetIndex, visible),
            "set_column_width" => SetColumnWidth(ctx, outputPath, sheetIndex, columnIndex, width),
            "set_row_height" => SetRowHeight(ctx, outputPath, sheetIndex, rowIndex, height),
            "set_background" => SetBackground(ctx, outputPath, sheetIndex, imagePath, removeBackground),
            "set_tab_color" => SetTabColor(ctx, outputPath, sheetIndex, color),
            "set_all" => SetAll(ctx, outputPath, sheetIndex, zoom, showGridlines, showRowColumnHeaders, showZeroValues,
                displayRightToLeft),
            "freeze_panes" => FreezePanes(ctx, outputPath, sheetIndex, freezeRow, freezeColumn, unfreeze),
            "split_window" => SplitWindow(ctx, outputPath, sheetIndex, splitRow, splitColumn, removeSplit),
            "auto_fit_column" => AutoFitColumn(ctx, outputPath, sheetIndex, columnIndex, startRow, endRow),
            "auto_fit_row" => AutoFitRow(ctx, outputPath, sheetIndex, rowIndex, startColumn, endColumn),
            "show_formulas" => ShowFormulas(ctx, outputPath, sheetIndex, visible),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets the zoom level for a worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="zoom">The zoom percentage (10-400).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when zoom is not between 10 and 400.</exception>
    private static string SetZoom(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int zoom)
    {
        if (zoom < 10 || zoom > 400)
            throw new ArgumentException("Zoom must be between 10 and 400");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.Zoom = zoom;

        ctx.Save(outputPath);
        return $"Zoom level set to {zoom}% for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets the visibility of gridlines.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="visible">Whether gridlines should be visible.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetGridlines(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, bool visible)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.IsGridlinesVisible = visible;

        ctx.Save(outputPath);
        return $"Gridlines visibility set to {(visible ? "visible" : "hidden")}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets the visibility of row and column headers.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="visible">Whether row and column headers should be visible.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetHeaders(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, bool visible)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.IsRowColumnHeadersVisible = visible;

        ctx.Save(outputPath);
        return
            $"RowColumnHeaders visibility set to {(visible ? "visible" : "hidden")}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets whether zero values are displayed.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="visible">Whether zero values should be displayed.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetZeroValues(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, bool visible)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.DisplayZeros = visible;

        ctx.Save(outputPath);
        return $"Zero values visibility set to {(visible ? "visible" : "hidden")}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets the width of a column.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="columnIndex">The zero-based column index.</param>
    /// <param name="width">The column width in characters.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetColumnWidth(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int columnIndex, double width)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.SetColumnWidth(columnIndex, width);

        ctx.Save(outputPath);
        return $"Column {columnIndex} width set to {width} characters. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets the height of a row.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="rowIndex">The zero-based row index.</param>
    /// <param name="height">The row height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string SetRowHeight(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int rowIndex,
        double height)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        worksheet.Cells.SetRowHeight(rowIndex, height);

        ctx.Save(outputPath);
        return $"Row {rowIndex} height set to {height} points. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets or removes the background image for a worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="imagePath">The path to the background image file.</param>
    /// <param name="removeBackground">Whether to remove the background image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    /// <exception cref="ArgumentException">Thrown when neither imagePath nor removeBackground is provided.</exception>
    private static string SetBackground(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? imagePath, bool removeBackground)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (removeBackground)
        {
            worksheet.BackgroundImage = null;
        }
        else if (!string.IsNullOrEmpty(imagePath))
        {
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            var imageBytes = File.ReadAllBytes(imagePath);
            worksheet.BackgroundImage = imageBytes;
        }
        else
        {
            throw new ArgumentException("Either imagePath or removeBackground must be provided");
        }

        ctx.Save(outputPath);
        return removeBackground
            ? $"Background image removed from sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}"
            : $"Background image set for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets the tab color for a worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="color">The color in hex format (e.g., '#FF0000').</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when color is null or empty.</exception>
    private static string SetTabColor(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? color)
    {
        if (string.IsNullOrEmpty(color))
            throw new ArgumentException("color is required for set_tab_color operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var parsedColor = ColorHelper.ParseColor(color);
        worksheet.TabColor = parsedColor;

        ctx.Save(outputPath);
        return $"Sheet tab color set to {color}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets multiple view settings at once.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="zoom">The zoom percentage (10-400).</param>
    /// <param name="showGridlines">Whether to show gridlines.</param>
    /// <param name="showRowColumnHeaders">Whether to show row and column headers.</param>
    /// <param name="showZeroValues">Whether to show zero values.</param>
    /// <param name="displayRightToLeft">Whether to display right to left.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when zoom is not between 10 and 400.</exception>
    private static string SetAll(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int zoom,
        bool? showGridlines, bool? showRowColumnHeaders, bool? showZeroValues, bool? displayRightToLeft)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (zoom != 100)
        {
            if (zoom < 10 || zoom > 400)
                throw new ArgumentException("Zoom must be between 10 and 400");
            worksheet.Zoom = zoom;
        }

        if (showGridlines.HasValue)
            worksheet.IsGridlinesVisible = showGridlines.Value;

        if (showRowColumnHeaders.HasValue)
            worksheet.IsRowColumnHeadersVisible = showRowColumnHeaders.Value;

        if (showZeroValues.HasValue)
            worksheet.DisplayZeros = showZeroValues.Value;

        if (displayRightToLeft.HasValue)
            worksheet.DisplayRightToLeft = displayRightToLeft.Value;

        ctx.Save(outputPath);
        return $"View settings updated for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Freezes or unfreezes panes in a worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="freezeRow">The row index to freeze at (0-based).</param>
    /// <param name="freezeColumn">The column index to freeze at (0-based).</param>
    /// <param name="unfreeze">Whether to unfreeze panes.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither freezeRow, freezeColumn, nor unfreeze is provided.</exception>
    private static string FreezePanes(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? freezeRow, int? freezeColumn, bool unfreeze)
    {
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);
        return unfreeze
            ? $"Panes unfrozen for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}"
            : $"Panes frozen at row {freezeRow ?? 0}, column {freezeColumn ?? 0} for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Splits or unsplits the window in a worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="splitRow">The row position to split at in pixels.</param>
    /// <param name="splitColumn">The column position to split at in pixels.</param>
    /// <param name="removeSplit">Whether to remove the window split.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither splitRow, splitColumn, nor removeSplit is provided.</exception>
    private static string SplitWindow(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? splitRow, int? splitColumn, bool removeSplit)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (removeSplit)
        {
            worksheet.RemoveSplit();
        }
        else if (splitRow.HasValue || splitColumn.HasValue)
        {
            var row = splitRow ?? 0;
            var col = splitColumn ?? 0;
            worksheet.ActiveCell = CellsHelper.CellIndexToName(row, col);
            worksheet.Split();
        }
        else
        {
            throw new ArgumentException("Either splitRow, splitColumn, or removeSplit must be provided");
        }

        ctx.Save(outputPath);
        return removeSplit
            ? $"Window split removed for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}"
            : $"Window split at row {splitRow ?? 0}, column {splitColumn ?? 0} for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Auto-fits a column width to its content.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="columnIndex">The zero-based column index to auto-fit.</param>
    /// <param name="startRow">The start row index for auto fit range (0-based).</param>
    /// <param name="endRow">The end row index for auto fit range (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string AutoFitColumn(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int columnIndex, int? startRow, int? endRow)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (startRow.HasValue && endRow.HasValue)
            worksheet.AutoFitColumn(columnIndex, startRow.Value, endRow.Value);
        else
            worksheet.AutoFitColumn(columnIndex);

        ctx.Save(outputPath);
        return $"Column {columnIndex} auto-fitted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Auto-fits a row height to its content.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="rowIndex">The zero-based row index to auto-fit.</param>
    /// <param name="startColumn">The start column index for auto fit range (0-based).</param>
    /// <param name="endColumn">The end column index for auto fit range (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string AutoFitRow(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int rowIndex, int? startColumn, int? endColumn)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (startColumn.HasValue && endColumn.HasValue)
            worksheet.AutoFitRow(rowIndex, startColumn.Value, endColumn.Value);
        else
            worksheet.AutoFitRow(rowIndex);

        ctx.Save(outputPath);
        return $"Row {rowIndex} auto-fitted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Sets whether formulas are displayed instead of values.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="visible">Whether formulas should be displayed instead of values.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string ShowFormulas(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, bool visible)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        worksheet.ShowFormulas = visible;

        ctx.Save(outputPath);
        return $"Formulas {(visible ? "shown" : "hidden")} for sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}