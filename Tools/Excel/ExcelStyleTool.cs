using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel styles (create, apply, get, format cells)
///     Merges: ExcelCreateStyleTool, ExcelApplyStyleTool, ExcelGetStylesTool, ExcelFormatCellsTool,
///     ExcelBatchFormatCellsTool, ExcelGetCellFormatTool, ExcelCopySheetFormatTool
/// </summary>
[McpServerToolType]
public class ExcelStyleTool
{
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
    [McpServerTool(Name = "excel_style")]
    [Description(@"Manage Excel styles. Supports 3 operations: format, get_format, copy_sheet_format.

Usage examples:
- Format cells: excel_style(operation='format', path='book.xlsx', range='A1:B10', fontName='Arial', fontSize=12, bold=true)
- Pattern fill: excel_style(operation='format', path='book.xlsx', range='A1', patternType='DiagonalStripe', backgroundColor='#FF0000', patternColor='#FFFFFF')
- Get format (full): excel_style(operation='get_format', path='book.xlsx', range='A1')
- Get format (simple): excel_style(operation='get_format', path='book.xlsx', range='A1', fields='font,color')
- Copy sheet format: excel_style(operation='copy_sheet_format', path='book.xlsx', sourceSheetIndex=0, targetSheetIndex=1)")]
    public string Execute(
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

        return operation.ToLower() switch
        {
            "format" => FormatCells(ctx, outputPath, sheetIndex, range, ranges, fontName, fontSize, bold, italic,
                fontColor, backgroundColor, patternType, patternColor, numberFormat, borderStyle, borderColor,
                horizontalAlignment, verticalAlignment),
            "get_format" => GetCellFormat(ctx, sheetIndex, cell, range, fields),
            "copy_sheet_format" => CopySheetFormat(ctx, outputPath, sourceSheetIndex, targetSheetIndex,
                copyColumnWidths, copyRowHeights),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Applies formatting to cells in a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to format.</param>
    /// <param name="rangesJson">JSON array of cell ranges for batch formatting.</param>
    /// <param name="fontName">The font name to apply.</param>
    /// <param name="fontSize">The font size to apply.</param>
    /// <param name="bold">Whether to apply bold formatting.</param>
    /// <param name="italic">Whether to apply italic formatting.</param>
    /// <param name="fontColor">The font color in hex format.</param>
    /// <param name="backgroundColor">The background color in hex format.</param>
    /// <param name="patternType">The fill pattern type.</param>
    /// <param name="patternColor">The pattern color in hex format.</param>
    /// <param name="numberFormat">The number format string.</param>
    /// <param name="borderStyle">The border style.</param>
    /// <param name="borderColor">The border color in hex format.</param>
    /// <param name="horizontalAlignment">The horizontal alignment.</param>
    /// <param name="verticalAlignment">The vertical alignment.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the font color is invalid or neither range nor ranges is provided.</exception>
    private static string FormatCells(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? range, string? rangesJson, string? fontName, int? fontSize, bool? bold, bool? italic,
        string? fontColor, string? backgroundColor, string? patternType, string? patternColor,
        string? numberFormat, string? borderStyle, string? borderColor,
        string? horizontalAlignment, string? verticalAlignment)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var style = workbook.CreateStyle();

        try
        {
            FontHelper.Excel.ApplyFontSettings(
                style,
                fontName,
                fontSize,
                bold,
                italic,
                fontColor
            );
        }
        catch (Exception colorEx) when (colorEx is ArgumentException && !string.IsNullOrWhiteSpace(fontColor))
        {
            throw new ArgumentException(
                $"Unable to parse font color '{fontColor}': {colorEx.Message}. Please use a valid color format (e.g., #FF0000, 255,0,0, or red)");
        }

        if (!string.IsNullOrWhiteSpace(backgroundColor) || !string.IsNullOrWhiteSpace(patternType))
        {
            var bgPattern = BackgroundType.Solid;
            if (!string.IsNullOrWhiteSpace(patternType))
                bgPattern = patternType.ToLower() switch
                {
                    "solid" => BackgroundType.Solid,
                    "gray50" => BackgroundType.Gray50,
                    "gray75" => BackgroundType.Gray75,
                    "gray25" => BackgroundType.Gray25,
                    "horizontalstripe" => BackgroundType.HorizontalStripe,
                    "verticalstripe" => BackgroundType.VerticalStripe,
                    "diagonalstripe" => BackgroundType.DiagonalStripe,
                    "reversediagonalstripe" => BackgroundType.ReverseDiagonalStripe,
                    "diagonalcrosshatch" => BackgroundType.DiagonalCrosshatch,
                    "thickdiagonalcrosshatch" => BackgroundType.ThickDiagonalCrosshatch,
                    "thinhorizontalstripe" => BackgroundType.ThinHorizontalStripe,
                    "thinverticalstripe" => BackgroundType.ThinVerticalStripe,
                    "thinreversediagonalstripe" => BackgroundType.ThinReverseDiagonalStripe,
                    "thindiagonalstripe" => BackgroundType.ThinDiagonalStripe,
                    "thinhorizontalcrosshatch" => BackgroundType.ThinHorizontalCrosshatch,
                    "thindiagonalcrosshatch" => BackgroundType.ThinDiagonalCrosshatch,
                    _ => BackgroundType.Solid
                };

            style.Pattern = bgPattern;

            if (!string.IsNullOrWhiteSpace(backgroundColor))
                style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, true);

            if (!string.IsNullOrWhiteSpace(patternColor))
                style.BackgroundColor = ColorHelper.ParseColor(patternColor, true);
        }

        if (!string.IsNullOrEmpty(numberFormat))
        {
            if (int.TryParse(numberFormat, out var formatNumber))
                style.Number = formatNumber;
            else
                style.Custom = numberFormat;
        }

        if (!string.IsNullOrEmpty(horizontalAlignment))
            style.HorizontalAlignment = horizontalAlignment.ToLower() switch
            {
                "left" => TextAlignmentType.Left,
                "center" => TextAlignmentType.Center,
                "right" => TextAlignmentType.Right,
                _ => TextAlignmentType.Left
            };

        if (!string.IsNullOrEmpty(verticalAlignment))
            style.VerticalAlignment = verticalAlignment.ToLower() switch
            {
                "top" => TextAlignmentType.Top,
                "center" => TextAlignmentType.Center,
                "bottom" => TextAlignmentType.Bottom,
                _ => TextAlignmentType.Center
            };

        if (!string.IsNullOrEmpty(borderStyle))
        {
            var borderType = borderStyle.ToLower() switch
            {
                "none" => CellBorderType.None,
                "thin" => CellBorderType.Thin,
                "medium" => CellBorderType.Medium,
                "thick" => CellBorderType.Thick,
                "dotted" => CellBorderType.Dotted,
                "dashed" => CellBorderType.Dashed,
                "double" => CellBorderType.Double,
                _ => CellBorderType.Thin
            };

            var borderColorValue = Color.Black;
            if (!string.IsNullOrWhiteSpace(borderColor))
                borderColorValue = ColorHelper.ParseColor(borderColor, true);

            style.SetBorder(BorderType.TopBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.BottomBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.LeftBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.RightBorder, borderType, borderColorValue);
        }

        var styleFlag = new StyleFlag
        {
            All = true,
            Borders = !string.IsNullOrEmpty(borderStyle)
        };

        if (!string.IsNullOrEmpty(rangesJson))
        {
            var rangesList = JsonSerializer.Deserialize<List<string>>(rangesJson);
            if (rangesList != null)
                foreach (var rangeStr in rangesList)
                    if (!string.IsNullOrEmpty(rangeStr))
                    {
                        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, rangeStr);
                        cellRange.ApplyStyle(style, styleFlag);
                    }
        }
        else if (!string.IsNullOrEmpty(range))
        {
            var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
            cellRange.ApplyStyle(style, styleFlag);
        }
        else
        {
            throw new ArgumentException("Either range or ranges must be provided for format operation");
        }

        ctx.Save(outputPath);
        return $"Cells formatted in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets format information from cells.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address.</param>
    /// <param name="range">The cell range.</param>
    /// <param name="fieldsParam">Comma-separated list of fields to retrieve.</param>
    /// <returns>A JSON string containing the format information.</returns>
    /// <exception cref="ArgumentException">Thrown when neither cell nor range is provided, or the cell range is invalid.</exception>
    private static string GetCellFormat(DocumentContext<Workbook> ctx, int sheetIndex, string? cell, string? range,
        string? fieldsParam)
    {
        if (string.IsNullOrEmpty(cell) && string.IsNullOrEmpty(range))
            throw new ArgumentException("Either cell or range is required for get_format operation");

        var cellOrRange = cell ?? range!;
        var requestedFields = ParseFields(fieldsParam);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        try
        {
            var cellRange = ExcelHelper.CreateRange(cells, cellOrRange);
            var startRow = cellRange.FirstRow;
            var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
            var startCol = cellRange.FirstColumn;
            var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

            List<Dictionary<string, object?>> cellList = [];
            for (var row = startRow; row <= endRow; row++)
            for (var col = startCol; col <= endCol; col++)
            {
                var cellObj = cells[row, col];
                var style = cellObj.GetStyle();

                var cellData = new Dictionary<string, object?>
                {
                    ["cell"] = CellsHelper.CellIndexToName(row, col)
                };

                if (requestedFields.Contains("value") || requestedFields.Contains("all"))
                {
                    cellData["value"] = cellObj.Value?.ToString() ?? "(empty)";
                    cellData["formula"] = cellObj.Formula;
                    cellData["dataType"] = cellObj.Type.ToString();
                }

                var formatData = new Dictionary<string, object?>();

                if (requestedFields.Contains("font") || requestedFields.Contains("all"))
                {
                    formatData["fontName"] = style.Font.Name;
                    formatData["fontSize"] = style.Font.Size;
                    formatData["bold"] = style.Font.IsBold;
                    formatData["italic"] = style.Font.IsItalic;
                    formatData["underline"] = style.Font.Underline.ToString();
                    formatData["strikethrough"] = style.Font.IsStrikeout;
                }

                if (requestedFields.Contains("color") || requestedFields.Contains("all"))
                {
                    formatData["fontColor"] = style.Font.Color.ToString();
                    formatData["foregroundColor"] = style.ForegroundColor.ToString();
                    formatData["backgroundColor"] = style.BackgroundColor.ToString();
                    formatData["patternType"] = style.Pattern.ToString();
                }

                if (requestedFields.Contains("alignment") || requestedFields.Contains("all"))
                {
                    formatData["horizontalAlignment"] = style.HorizontalAlignment.ToString();
                    formatData["verticalAlignment"] = style.VerticalAlignment.ToString();
                }

                if (requestedFields.Contains("number") || requestedFields.Contains("all"))
                {
                    formatData["numberFormat"] = style.Number;
                    formatData["customFormat"] = style.Custom;
                }

                if (requestedFields.Contains("border") || requestedFields.Contains("all"))
                {
                    var topBorder = style.Borders[BorderType.TopBorder];
                    var bottomBorder = style.Borders[BorderType.BottomBorder];
                    var leftBorder = style.Borders[BorderType.LeftBorder];
                    var rightBorder = style.Borders[BorderType.RightBorder];

                    formatData["borders"] = new
                    {
                        top = new { lineStyle = topBorder.LineStyle.ToString(), color = topBorder.Color.ToString() },
                        bottom = new
                            { lineStyle = bottomBorder.LineStyle.ToString(), color = bottomBorder.Color.ToString() },
                        left = new { lineStyle = leftBorder.LineStyle.ToString(), color = leftBorder.Color.ToString() },
                        right = new
                        {
                            lineStyle = rightBorder.LineStyle.ToString(), color = rightBorder.Color.ToString()
                        }
                    };
                }

                if (formatData.Count > 0)
                    cellData["format"] = formatData;

                cellList.Add(cellData);
            }

            var result = new
            {
                count = cellList.Count,
                worksheetName = worksheet.Name,
                range = cellOrRange,
                fields = fieldsParam ?? "all",
                items = cellList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Invalid cell range: '{cellOrRange}'. Expected format: single cell (e.g., 'A1') or range (e.g., 'A1:C5'). Error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Parses comma-separated field names into a set.
    /// </summary>
    /// <param name="fieldsParam">Comma-separated list of field names.</param>
    /// <returns>A set of field names, or a set containing "all" if no fields are specified.</returns>
    private static HashSet<string> ParseFields(string? fieldsParam)
    {
        if (string.IsNullOrWhiteSpace(fieldsParam))
            return ["all"];

        var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fieldsParam.Split(',',
                     StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            fields.Add(field.ToLower());

        return fields.Count == 0 ? ["all"] : fields;
    }

    /// <summary>
    ///     Copies format from one sheet to another.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sourceSheetIndex">The source worksheet index.</param>
    /// <param name="targetSheetIndex">The target worksheet index.</param>
    /// <param name="copyColumnWidths">Whether to copy column widths.</param>
    /// <param name="copyRowHeights">Whether to copy row heights.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string CopySheetFormat(DocumentContext<Workbook> ctx, string? outputPath,
        int sourceSheetIndex, int targetSheetIndex, bool copyColumnWidths, bool copyRowHeights)
    {
        var workbook = ctx.Document;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
        var targetSheet = ExcelHelper.GetWorksheet(workbook, targetSheetIndex);

        if (copyColumnWidths)
        {
            // Use MaxColumn to include columns with formatting but no data
            // If both are negative (no data/formatting), copy at least the first column
            var maxCol = Math.Max(sourceSheet.Cells.MaxColumn, sourceSheet.Cells.MaxDataColumn);
            maxCol = Math.Max(maxCol, 0); // Ensure at least column 0 is copied
            for (var i = 0; i <= maxCol; i++)
                targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));
        }

        if (copyRowHeights)
        {
            // Use MaxRow to include rows with formatting but no data
            // If both are negative (no data/formatting), copy at least the first row
            var maxRow = Math.Max(sourceSheet.Cells.MaxRow, sourceSheet.Cells.MaxDataRow);
            maxRow = Math.Max(maxRow, 0); // Ensure at least row 0 is copied
            for (var i = 0; i <= maxRow; i++)
                targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));
        }

        ctx.Save(outputPath);
        return
            $"Sheet format copied from sheet {sourceSheetIndex} to sheet {targetSheetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}