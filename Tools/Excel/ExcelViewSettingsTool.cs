using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel view settings (zoom, gridlines, headers, zero values, etc.)
///     Merges: ExcelSetZoomTool, ExcelSetGridlinesVisibleTool, ExcelSetRowColumnHeadersVisibleTool,
///     ExcelSetZeroValuesVisibleTool, ExcelSetViewSettingsTool, ExcelSetColumnWidthTool, ExcelSetRowHeightTool,
///     ExcelSetSheetBackgroundTool, ExcelSetSheetTabColorTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.ViewSettings")]
[McpServerToolType]
public class ExcelViewSettingsTool
{
    /// <summary>
    ///     Handler registry for view settings operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Server configuration for path-allowlist enforcement.
    /// </summary>
    private readonly ServerConfig? _serverConfig;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelViewSettingsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    /// <param name="serverConfig">Optional server config for path allowlist enforcement.</param>
    public ExcelViewSettingsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null,
        ServerConfig? serverConfig = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _serverConfig = serverConfig;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.ViewSettings");
    }

    /// <summary>
    ///     Executes an Excel view settings operation (set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width,
    ///     set_row_height, set_background, set_tab, set, freeze, split, autofit_col, autofit_row,
    ///     show_formulas).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set_zoom, set_gridlines, set_headers, set_zero_values,
    ///     set_column_width, set_row_height, set_background, set_tab, set, freeze, split,
    ///     autofit_col, autofit_row, show_formulas.
    /// </param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="zoom">Zoom percentage (10-400, optional for set_zoom, default: 100).</param>
    /// <param name="visible">Visibility (optional for set_gridlines/set_headers/set_zero_values/show_formulas, default: true).</param>
    /// <param name="columnIndex">Column index (0-based, optional for set_column_width/autofit_col, default: 0).</param>
    /// <param name="width">Column width in characters (optional for set_column_width, default: 8.43, the Excel default width).</param>
    /// <param name="rowIndex">Row index (0-based, optional for set_row_height/autofit_row, default: 0).</param>
    /// <param name="height">Row height in points (optional for set_row_height, default: 15, the Excel default height).</param>
    /// <param name="imagePath">Background image file path (required for set_background).</param>
    /// <param name="removeBackground">Remove background image (for set_background).</param>
    /// <param name="color">Color in hex format (e.g., '#FF0000', required for set_tab).</param>
    /// <param name="showGridlines">Show gridlines (for set).</param>
    /// <param name="showRowColumnHeaders">Show row/column headers (for set).</param>
    /// <param name="showZeroValues">Show zero values (for set).</param>
    /// <param name="displayRightToLeft">Display right to left (for set).</param>
    /// <param name="freezeRow">Row index to freeze at (0-based, for freeze).</param>
    /// <param name="freezeColumn">Column index to freeze at (0-based, for freeze).</param>
    /// <param name="unfreeze">Remove frozen panes (for freeze).</param>
    /// <param name="splitRow">Row position to split at in pixels (for split).</param>
    /// <param name="splitColumn">Column position to split at in pixels (for split).</param>
    /// <param name="removeSplit">Remove window split (for split).</param>
    /// <param name="startRow">Start row index for auto fit range (0-based, for autofit_col).</param>
    /// <param name="endRow">End row index for auto fit range (0-based, for autofit_col).</param>
    /// <param name="startColumn">Start column index for auto fit range (0-based, for autofit_row).</param>
    /// <param name="endColumn">End column index for auto fit range (0-based, for autofit_row).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_view_settings",
        Title = "Excel View Settings Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage Excel view settings. Supports 14 operations: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab, set, freeze, split, autofit_col, autofit_row, show_formulas.

Usage examples:
- Set zoom: excel_view_settings(operation='set_zoom', path='book.xlsx', zoom=150)
- Set gridlines: excel_view_settings(operation='set_gridlines', path='book.xlsx', visible=false)
- Set column width: excel_view_settings(operation='set_column_width', path='book.xlsx', columnIndex=0, width=20)
- Set row height: excel_view_settings(operation='set_row_height', path='book.xlsx', rowIndex=0, height=30)
- Freeze panes: excel_view_settings(operation='freeze', path='book.xlsx', freezeRow=1, freezeColumn=1)
- Split window: excel_view_settings(operation='split', path='book.xlsx', splitRow=5, splitColumn=2)
- Auto fit column: excel_view_settings(operation='autofit_col', path='book.xlsx', columnIndex=0)
- Show formulas: excel_view_settings(operation='show_formulas', path='book.xlsx', visible=true)
- Set all: excel_view_settings(operation='set', path='book.xlsx', zoom=150, showGridlines=true)")]
    public object Execute(
        [Description(
            "Operation: set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab, set, freeze, split, autofit_col, autofit_row, show_formulas")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Zoom percentage (10-400, optional for set_zoom, default: 100)")]
        int zoom = 100,
        [Description(
            "Visibility (optional for set_gridlines/set_headers/set_zero_values/show_formulas, default: true)")]
        bool visible = true,
        [Description("Column index (0-based, optional for set_column_width/autofit_col, default: 0)")]
        int columnIndex = 0,
        [Description(
            "Column width in characters (optional for set_column_width, default: 8.43, the Excel default width)")]
        double width = 8.43,
        [Description("Row index (0-based, optional for set_row_height/autofit_row, default: 0)")]
        int rowIndex = 0,
        [Description("Row height in points (optional for set_row_height, default: 15, the Excel default height)")]
        double height = 15,
        [Description("Background image file path (required for set_background)")]
        string? imagePath = null,
        [Description("Remove background image (for set_background)")]
        bool removeBackground = false,
        [Description("Color in hex format (e.g., '#FF0000', required for set_tab)")]
        string? color = null,
        [Description("Show gridlines (for set)")]
        bool? showGridlines = null,
        [Description("Show row/column headers (for set)")]
        bool? showRowColumnHeaders = null,
        [Description("Show zero values (for set)")]
        bool? showZeroValues = null,
        [Description("Display right to left (for set)")]
        bool? displayRightToLeft = null,
        [Description("Row index to freeze at (0-based, for freeze)")]
        int? freezeRow = null,
        [Description("Column index to freeze at (0-based, for freeze)")]
        int? freezeColumn = null,
        [Description("Remove frozen panes (for freeze)")]
        bool unfreeze = false,
        [Description("Row position to split at in pixels (for split)")]
        int? splitRow = null,
        [Description("Column position to split at in pixels (for split)")]
        int? splitColumn = null,
        [Description("Remove window split (for split)")]
        bool removeSplit = false,
        [Description("Start row index for auto fit range (0-based, for autofit_col)")]
        int? startRow = null,
        [Description("End row index for auto fit range (0-based, for autofit_col)")]
        int? endRow = null,
        [Description("Start column index for auto fit range (0-based, for autofit_row)")]
        int? startColumn = null,
        [Description("End column index for auto fit range (0-based, for autofit_row)")]
        int? endColumn = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor,
            serverConfig: _serverConfig);

        var parameters = BuildParameters(operation, sheetIndex, zoom, visible, columnIndex, width, rowIndex, height,
            imagePath, removeBackground, color, showGridlines, showRowColumnHeaders, showZeroValues, displayRightToLeft,
            freezeRow, freezeColumn, unfreeze, splitRow, splitColumn, removeSplit, startRow, endRow, startColumn,
            endColumn);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath,
            ServerConfig = _serverConfig
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        int zoom,
        bool visible,
        int columnIndex,
        double width,
        int rowIndex,
        double height,
        string? imagePath,
        bool removeBackground,
        string? color,
        bool? showGridlines,
        bool? showRowColumnHeaders,
        bool? showZeroValues,
        bool? displayRightToLeft,
        int? freezeRow,
        int? freezeColumn,
        bool unfreeze,
        int? splitRow,
        int? splitColumn,
        bool removeSplit,
        int? startRow,
        int? endRow,
        int? startColumn,
        int? endColumn)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "set_zoom" => BuildSetZoomParameters(parameters, zoom),
            "set_gridlines" or "set_headers" or "set_zero_values" or "show_formulas"
                => BuildVisibilityParameters(parameters, visible),
            "set_column_width" => BuildColumnWidthParameters(parameters, columnIndex, width),
            "set_row_height" => BuildRowHeightParameters(parameters, rowIndex, height),
            "set_background" => BuildBackgroundParameters(parameters, imagePath, removeBackground),
            "set_tab" => BuildTabColorParameters(parameters, color),
            "set" => BuildSetAllParameters(parameters, zoom, showGridlines, showRowColumnHeaders,
                showZeroValues, displayRightToLeft),
            "freeze" => BuildFreezePanesParameters(parameters, freezeRow, freezeColumn, unfreeze),
            "split" => BuildSplitWindowParameters(parameters, splitRow, splitColumn, removeSplit),
            "autofit_col" => BuildAutoFitColumnParameters(parameters, columnIndex, startRow, endRow),
            "autofit_row" => BuildAutoFitRowParameters(parameters, rowIndex, startColumn, endColumn),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the set zoom operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="zoom">The zoom percentage value (10-400).</param>
    /// <returns>OperationParameters configured for the set zoom operation.</returns>
    private static OperationParameters BuildSetZoomParameters(OperationParameters parameters, int zoom)
    {
        parameters.Set("zoom", zoom);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for visibility-related operations (gridlines, headers, zero values, show formulas).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="visible">Whether the element should be visible.</param>
    /// <returns>OperationParameters configured for the visibility operation.</returns>
    private static OperationParameters BuildVisibilityParameters(OperationParameters parameters, bool visible)
    {
        parameters.Set("visible", visible);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set column width operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="columnIndex">The column index (0-based).</param>
    /// <param name="width">The column width in characters.</param>
    /// <returns>OperationParameters configured for the set column width operation.</returns>
    private static OperationParameters BuildColumnWidthParameters(OperationParameters parameters, int columnIndex,
        double width)
    {
        parameters.Set("columnIndex", columnIndex);
        parameters.Set("width", width);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set row height operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="rowIndex">The row index (0-based).</param>
    /// <param name="height">The row height in points.</param>
    /// <returns>OperationParameters configured for the set row height operation.</returns>
    private static OperationParameters BuildRowHeightParameters(OperationParameters parameters, int rowIndex,
        double height)
    {
        parameters.Set("rowIndex", rowIndex);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set background operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="imagePath">The path to the background image file.</param>
    /// <param name="removeBackground">Whether to remove the background image.</param>
    /// <returns>OperationParameters configured for the set background operation.</returns>
    private static OperationParameters BuildBackgroundParameters(OperationParameters parameters, string? imagePath,
        bool removeBackground)
    {
        if (imagePath != null) parameters.Set("imagePath", imagePath);
        parameters.Set("removeBackground", removeBackground);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set tab color operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="color">The tab color in hex format (e.g., '#FF0000').</param>
    /// <returns>OperationParameters configured for the set tab color operation.</returns>
    private static OperationParameters BuildTabColorParameters(OperationParameters parameters, string? color)
    {
        if (color != null) parameters.Set("color", color);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set all view settings operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="zoom">The zoom percentage value.</param>
    /// <param name="showGridlines">Whether to show gridlines.</param>
    /// <param name="showRowColumnHeaders">Whether to show row/column headers.</param>
    /// <param name="showZeroValues">Whether to show zero values.</param>
    /// <param name="displayRightToLeft">Whether to display right to left.</param>
    /// <returns>OperationParameters configured for the set all operation.</returns>
    private static OperationParameters BuildSetAllParameters(OperationParameters parameters, int zoom,
        bool? showGridlines, bool? showRowColumnHeaders, bool? showZeroValues, bool? displayRightToLeft)
    {
        parameters.Set("zoom", zoom);
        if (showGridlines.HasValue) parameters.Set("showGridlines", showGridlines.Value);
        if (showRowColumnHeaders.HasValue) parameters.Set("showRowColumnHeaders", showRowColumnHeaders.Value);
        if (showZeroValues.HasValue) parameters.Set("showZeroValues", showZeroValues.Value);
        if (displayRightToLeft.HasValue) parameters.Set("displayRightToLeft", displayRightToLeft.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the freeze panes operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="freezeRow">The row index to freeze at (0-based).</param>
    /// <param name="freezeColumn">The column index to freeze at (0-based).</param>
    /// <param name="unfreeze">Whether to remove frozen panes.</param>
    /// <returns>OperationParameters configured for the freeze panes operation.</returns>
    private static OperationParameters BuildFreezePanesParameters(OperationParameters parameters, int? freezeRow,
        int? freezeColumn, bool unfreeze)
    {
        if (freezeRow.HasValue) parameters.Set("freezeRow", freezeRow.Value);
        if (freezeColumn.HasValue) parameters.Set("freezeColumn", freezeColumn.Value);
        parameters.Set("unfreeze", unfreeze);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the split window operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="splitRow">The row position to split at in pixels.</param>
    /// <param name="splitColumn">The column position to split at in pixels.</param>
    /// <param name="removeSplit">Whether to remove the window split.</param>
    /// <returns>OperationParameters configured for the split window operation.</returns>
    private static OperationParameters BuildSplitWindowParameters(OperationParameters parameters, int? splitRow,
        int? splitColumn, bool removeSplit)
    {
        if (splitRow.HasValue) parameters.Set("splitRow", splitRow.Value);
        if (splitColumn.HasValue) parameters.Set("splitColumn", splitColumn.Value);
        parameters.Set("removeSplit", removeSplit);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the auto fit column operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="columnIndex">The column index to auto fit (0-based).</param>
    /// <param name="startRow">The start row index for auto fit range (0-based).</param>
    /// <param name="endRow">The end row index for auto fit range (0-based).</param>
    /// <returns>OperationParameters configured for the auto fit column operation.</returns>
    private static OperationParameters BuildAutoFitColumnParameters(OperationParameters parameters, int columnIndex,
        int? startRow, int? endRow)
    {
        parameters.Set("columnIndex", columnIndex);
        if (startRow.HasValue) parameters.Set("startRow", startRow.Value);
        if (endRow.HasValue) parameters.Set("endRow", endRow.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the auto fit row operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="rowIndex">The row index to auto fit (0-based).</param>
    /// <param name="startColumn">The start column index for auto fit range (0-based).</param>
    /// <param name="endColumn">The end column index for auto fit range (0-based).</param>
    /// <returns>OperationParameters configured for the auto fit row operation.</returns>
    private static OperationParameters BuildAutoFitRowParameters(OperationParameters parameters, int rowIndex,
        int? startColumn, int? endColumn)
    {
        parameters.Set("rowIndex", rowIndex);
        if (startColumn.HasValue) parameters.Set("startColumn", startColumn.Value);
        if (endColumn.HasValue) parameters.Set("endColumn", endColumn.Value);
        return parameters;
    }
}
