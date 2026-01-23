using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel styles (format, get_format, copy_sheet_format)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Style")]
[McpServerToolType]
public class ExcelStyleTool
{
    /// <summary>
    ///     Handler registry for style operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelStyleTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelStyleTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Style");
    }

    /// <summary>
    ///     Executes an Excel style operation (format, get_format, copy_sheet_format).
    /// </summary>
    /// <param name="operation">The operation to perform: format, get_format, copy_sheet_format.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="sourceSheetIndex">Source sheet index (0-based, required for copy_sheet_format).</param>
    /// <param name="targetSheetIndex">Target sheet index (0-based, required for copy_sheet_format).</param>
    /// <param name="range">Cell range (e.g., 'A1:C5', required for format).</param>
    /// <param name="cell">Cell address or range (e.g., 'A1' or 'A1:C5', for get_format).</param>
    /// <param name="fields">Comma-separated list of fields: font, color, alignment, border, number, value, all.</param>
    /// <param name="ranges">Array of cell ranges for batch format (JSON array string).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontSize">Font size.</param>
    /// <param name="bold">Bold.</param>
    /// <param name="italic">Italic.</param>
    /// <param name="fontColor">Font/text color (hex format like '#FF0000').</param>
    /// <param name="backgroundColor">Background/foreground color for fill (hex format).</param>
    /// <param name="patternType">Fill pattern type (Solid, Gray50, HorizontalStripe, etc.).</param>
    /// <param name="patternColor">Pattern/background color for two-color patterns (hex format).</param>
    /// <param name="numberFormat">Number format string (e.g., 'yyyy-mm-dd', '#,##0.00').</param>
    /// <param name="borderStyle">Border style (None, Thin, Medium, Thick).</param>
    /// <param name="borderColor">Border color (hex format).</param>
    /// <param name="horizontalAlignment">Horizontal alignment (Left, Center, Right).</param>
    /// <param name="verticalAlignment">Vertical alignment (Top, Center, Bottom).</param>
    /// <param name="copyColumnWidths">Copy column widths (default: true).</param>
    /// <param name="copyRowHeights">Copy row heights (default: true).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_format operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_style",
        Title = "Excel Style Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel styles. Supports 3 operations: format, get_format, copy_sheet_format.

Usage examples:
- Format cells: excel_style(operation='format', path='book.xlsx', range='A1:B10', fontName='Arial', fontSize=12, bold=true)
- Pattern fill: excel_style(operation='format', path='book.xlsx', range='A1', patternType='DiagonalStripe', backgroundColor='#FF0000', patternColor='#FFFFFF')
- Get format (full): excel_style(operation='get_format', path='book.xlsx', range='A1')
- Get format (simple): excel_style(operation='get_format', path='book.xlsx', range='A1', fields='font,color')
- Copy sheet format: excel_style(operation='copy_sheet_format', path='book.xlsx', sourceSheetIndex=0, targetSheetIndex=1)")]
    public object Execute(
        [Description("Operation to perform: format, get_format, copy_sheet_format")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Source sheet index (0-based, required for copy_sheet_format)")]
        int sourceSheetIndex = 0,
        [Description("Target sheet index (0-based, required for copy_sheet_format)")]
        int targetSheetIndex = 0,
        [Description("Cell range (e.g., 'A1:C5', required for format)")]
        string? range = null,
        [Description("Cell address or range (e.g., 'A1' or 'A1:C5', for get_format)")]
        string? cell = null,
        [Description("Comma-separated list of fields: font, color, alignment, border, number, value, all")]
        string? fields = null,
        [Description("Array of cell ranges for batch format (JSON array string)")]
        string? ranges = null,
        [Description("Font name")] string? fontName = null,
        [Description("Font size")] int? fontSize = null,
        [Description("Bold")] bool? bold = null,
        [Description("Italic")] bool? italic = null,
        [Description("Font/text color (hex format like '#FF0000')")]
        string? fontColor = null,
        [Description("Background/foreground color for fill (hex format)")]
        string? backgroundColor = null,
        [Description("Fill pattern type (Solid, Gray50, HorizontalStripe, etc.)")]
        string? patternType = null,
        [Description("Pattern/background color for two-color patterns (hex format)")]
        string? patternColor = null,
        [Description("Number format string (e.g., 'yyyy-mm-dd', '#,##0.00')")]
        string? numberFormat = null,
        [Description("Border style (None, Thin, Medium, Thick)")]
        string? borderStyle = null,
        [Description("Border color (hex format)")]
        string? borderColor = null,
        [Description("Horizontal alignment (Left, Center, Right)")]
        string? horizontalAlignment = null,
        [Description("Vertical alignment (Top, Center, Bottom)")]
        string? verticalAlignment = null,
        [Description("Copy column widths (default: true)")]
        bool copyColumnWidths = true,
        [Description("Copy row heights (default: true)")]
        bool copyRowHeights = true)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, sourceSheetIndex, targetSheetIndex, range, cell,
            fields, ranges, fontName, fontSize, bold, italic, fontColor, backgroundColor, patternType, patternColor,
            numberFormat, borderStyle, borderColor, horizontalAlignment, verticalAlignment,
            copyColumnWidths, copyRowHeights);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get_format", StringComparison.OrdinalIgnoreCase))
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

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
        int sourceSheetIndex,
        int targetSheetIndex,
        string? range,
        string? cell,
        string? fields,
        string? ranges,
        string? fontName,
        int? fontSize,
        bool? bold,
        bool? italic,
        string? fontColor,
        string? backgroundColor,
        string? patternType,
        string? patternColor,
        string? numberFormat,
        string? borderStyle,
        string? borderColor,
        string? horizontalAlignment,
        string? verticalAlignment,
        bool copyColumnWidths,
        bool copyRowHeights)
    {
        return operation.ToLowerInvariant() switch
        {
            "format" => BuildFormatParameters(sheetIndex, range, ranges, fontName, fontSize, bold, italic, fontColor,
                backgroundColor, patternType, patternColor, numberFormat, borderStyle, borderColor,
                horizontalAlignment, verticalAlignment),
            "get_format" => BuildGetFormatParameters(sheetIndex, cell, range, fields),
            "copy_sheet_format" => BuildCopySheetFormatParameters(sheetIndex, sourceSheetIndex, targetSheetIndex,
                copyColumnWidths, copyRowHeights),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the format cells operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="range">The cell range to format.</param>
    /// <param name="ranges">JSON array of cell ranges for batch formatting.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontSize">The font size.</param>
    /// <param name="bold">Whether to apply bold.</param>
    /// <param name="italic">Whether to apply italic.</param>
    /// <param name="fontColor">The font color (hex format).</param>
    /// <param name="backgroundColor">The background color (hex format).</param>
    /// <param name="patternType">The fill pattern type.</param>
    /// <param name="patternColor">The pattern color (hex format).</param>
    /// <param name="numberFormat">The number format string.</param>
    /// <param name="borderStyle">The border style.</param>
    /// <param name="borderColor">The border color (hex format).</param>
    /// <param name="horizontalAlignment">The horizontal alignment.</param>
    /// <param name="verticalAlignment">The vertical alignment.</param>
    /// <returns>OperationParameters configured for formatting cells.</returns>
    private static OperationParameters BuildFormatParameters(
        int sheetIndex, string? range, string? ranges,
        string? fontName, int? fontSize, bool? bold, bool? italic, string? fontColor, string? backgroundColor,
        string? patternType, string? patternColor, string? numberFormat, string? borderStyle, string? borderColor,
        string? horizontalAlignment, string? verticalAlignment)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        if (range != null) parameters.Set("range", range);
        if (ranges != null) parameters.Set("ranges", ranges);
        if (fontName != null) parameters.Set("fontName", fontName);
        if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
        if (bold.HasValue) parameters.Set("bold", bold.Value);
        if (italic.HasValue) parameters.Set("italic", italic.Value);
        if (fontColor != null) parameters.Set("fontColor", fontColor);
        if (backgroundColor != null) parameters.Set("backgroundColor", backgroundColor);
        if (patternType != null) parameters.Set("patternType", patternType);
        if (patternColor != null) parameters.Set("patternColor", patternColor);
        if (numberFormat != null) parameters.Set("numberFormat", numberFormat);
        if (borderStyle != null) parameters.Set("borderStyle", borderStyle);
        if (borderColor != null) parameters.Set("borderColor", borderColor);
        if (horizontalAlignment != null) parameters.Set("horizontalAlignment", horizontalAlignment);
        if (verticalAlignment != null) parameters.Set("verticalAlignment", verticalAlignment);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get format operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="range">The cell range.</param>
    /// <param name="fields">Comma-separated list of fields to retrieve.</param>
    /// <returns>OperationParameters configured for getting cell format.</returns>
    private static OperationParameters BuildGetFormatParameters(int sheetIndex, string? cell, string? range,
        string? fields)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        if (cell != null) parameters.Set("cell", cell);
        if (range != null) parameters.Set("range", range);
        if (fields != null) parameters.Set("fields", fields);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy sheet format operation.
    /// </summary>
    /// <param name="sheetIndex">The sheet index (0-based).</param>
    /// <param name="sourceSheetIndex">The source sheet index to copy format from.</param>
    /// <param name="targetSheetIndex">The target sheet index to copy format to.</param>
    /// <param name="copyColumnWidths">Whether to copy column widths.</param>
    /// <param name="copyRowHeights">Whether to copy row heights.</param>
    /// <returns>OperationParameters configured for copying sheet format.</returns>
    private static OperationParameters BuildCopySheetFormatParameters(int sheetIndex, int sourceSheetIndex,
        int targetSheetIndex, bool copyColumnWidths, bool copyRowHeights)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        parameters.Set("sourceSheetIndex", sourceSheetIndex);
        parameters.Set("targetSheetIndex", targetSheetIndex);
        parameters.Set("copyColumnWidths", copyColumnWidths);
        parameters.Set("copyRowHeights", copyRowHeights);
        return parameters;
    }
}
