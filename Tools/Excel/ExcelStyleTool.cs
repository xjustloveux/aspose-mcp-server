using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel styles (create, apply, get, format cells)
///     Merges: ExcelCreateStyleTool, ExcelApplyStyleTool, ExcelGetStylesTool, ExcelFormatCellsTool,
///     ExcelBatchFormatCellsTool, ExcelGetCellFormatTool, ExcelCopySheetFormatTool
/// </summary>
public class ExcelStyleTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel styles. Supports 3 operations: format, get_format, copy_sheet_format.

Usage examples:
- Format cells: excel_style(operation='format', path='book.xlsx', range='A1:B10', fontName='Arial', fontSize=12, bold=true)
- Pattern fill: excel_style(operation='format', path='book.xlsx', range='A1', patternType='DiagonalStripe', backgroundColor='#FF0000', patternColor='#FFFFFF')
- Get format (full): excel_style(operation='get_format', path='book.xlsx', range='A1')
- Get format (simple): excel_style(operation='get_format', path='book.xlsx', range='A1', fields='font,color')
- Copy sheet format: excel_style(operation='copy_sheet_format', path='book.xlsx', sourceSheetIndex=0, targetSheetIndex=1)";

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
                description = @"Operation to perform.
- 'format': Format cells (required params: path, range)
- 'get_format': Get cell format (required params: path, range)
- 'copy_sheet_format': Copy sheet format (required params: path, sourceSheetIndex, targetSheetIndex)",
                @enum = new[] { "format", "get_format", "copy_sheet_format" }
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
                description = "Source sheet index (0-based, required for copy_sheet_format)"
            },
            targetSheetIndex = new
            {
                type = "number",
                description = "Target sheet index (0-based, required for copy_sheet_format)"
            },
            range = new
            {
                type = "string",
                description =
                    "Cell range (e.g., 'A1:C5', required for format, optional for get_format as alternative to cell)"
            },
            cell = new
            {
                type = "string",
                description =
                    "Cell address or range (e.g., 'A1' or 'A1:C5', required for get_format, or use range as alternative)"
            },
            fields = new
            {
                type = "string",
                description =
                    "Comma-separated list of fields to return for get_format (optional, reduces token usage). Valid values: font, color, alignment, border, number, value, all. Default: all"
            },
            ranges = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of cell ranges (optional, for batch format)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic (optional)"
            },
            fontColor = new
            {
                type = "string",
                description = "Font/text color (hex format like '#FF0000', optional)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background/foreground color for fill (hex format like '#FFFF00', optional)"
            },
            patternType = new
            {
                type = "string",
                description =
                    "Fill pattern type (Solid, Gray50, Gray75, Gray25, HorizontalStripe, VerticalStripe, DiagonalStripe, ReverseDiagonalStripe, DiagonalCrosshatch, ThickDiagonalCrosshatch, ThinHorizontalStripe, ThinVerticalStripe, ThinReverseDiagonalStripe, ThinDiagonalStripe, ThinHorizontalCrosshatch, ThinDiagonalCrosshatch, optional, default: Solid)"
            },
            patternColor = new
            {
                type = "string",
                description =
                    "Pattern/background color for pattern fill (hex format, optional, used with patternType for two-color patterns)"
            },
            numberFormat = new
            {
                type = "string",
                description = "Number format string (e.g., 'yyyy-mm-dd', '#,##0.00', optional)"
            },
            borderStyle = new
            {
                type = "string",
                description = "Border style (None, Thin, Medium, Thick, optional)"
            },
            borderColor = new
            {
                type = "string",
                description = "Border color (hex format, optional)"
            },
            horizontalAlignment = new
            {
                type = "string",
                description = "Horizontal alignment (Left, Center, Right, optional)"
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment (Top, Center, Bottom, optional)"
            },
            copyColumnWidths = new
            {
                type = "boolean",
                description = "Copy column widths (optional, for copy_sheet_format, default: true)"
            },
            copyRowHeights = new
            {
                type = "boolean",
                description = "Copy row heights (optional, for copy_sheet_format, default: true)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for format/copy_sheet_format operations, defaults to input path)"
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
            "format" => await FormatCellsAsync(path, outputPath, sheetIndex, arguments),
            "get_format" => await GetCellFormatAsync(path, sheetIndex, arguments),
            "copy_sheet_format" => await CopySheetFormatAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Formats cells with specified style properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing range/ranges and various format properties</param>
    /// <returns>Success message with formatted sheet index</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when range/ranges is missing, or color format is invalid
    /// </exception>
    private Task<string> FormatCellsAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var range = ArgumentHelper.GetStringNullable(arguments, "range");
            var rangesArray = ArgumentHelper.GetArray(arguments, "ranges", false);
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontSize = ArgumentHelper.GetIntNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var fontColor = ArgumentHelper.GetStringNullable(arguments, "fontColor");
            var backgroundColor = ArgumentHelper.GetStringNullable(arguments, "backgroundColor");
            var patternType = ArgumentHelper.GetStringNullable(arguments, "patternType");
            var patternColor = ArgumentHelper.GetStringNullable(arguments, "patternColor");
            var numberFormat = ArgumentHelper.GetStringNullable(arguments, "numberFormat");
            var borderStyle = ArgumentHelper.GetStringNullable(arguments, "borderStyle");
            var borderColor = ArgumentHelper.GetStringNullable(arguments, "borderColor");
            var horizontalAlignment = ArgumentHelper.GetStringNullable(arguments, "horizontalAlignment");
            var verticalAlignment = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var style = workbook.CreateStyle();

            // Apply font settings using FontHelper
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
                // Re-throw color parsing errors with context
                throw new ArgumentException(
                    $"Unable to parse font color '{fontColor}': {colorEx.Message}. Please use a valid color format (e.g., #FF0000, 255,0,0, or red)");
            }

            if (!string.IsNullOrWhiteSpace(backgroundColor) || !string.IsNullOrWhiteSpace(patternType))
            {
                // Parse pattern type
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

                // Set foreground color (primary fill color)
                if (!string.IsNullOrWhiteSpace(backgroundColor))
                    style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, true);

                // Set background color (pattern color for two-color patterns)
                if (!string.IsNullOrWhiteSpace(patternColor))
                    style.BackgroundColor = ColorHelper.ParseColor(patternColor, true);
            }

            if (!string.IsNullOrEmpty(numberFormat))
            {
                // Try to parse as built-in format number, otherwise use Custom
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

            // Apply border settings
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
                    // Parse color with error handling - throws ArgumentException on failure
                    borderColorValue = ColorHelper.ParseColor(borderColor, true);

                // Set borders for all sides
                style.SetBorder(BorderType.TopBorder, borderType, borderColorValue);
                style.SetBorder(BorderType.BottomBorder, borderType, borderColorValue);
                style.SetBorder(BorderType.LeftBorder, borderType, borderColorValue);
                style.SetBorder(BorderType.RightBorder, borderType, borderColorValue);
            }

            // Create StyleFlag to specify which style properties to apply
            var styleFlag = new StyleFlag
            {
                All = true,
                Borders = !string.IsNullOrEmpty(borderStyle)
            };

            if (rangesArray is { Count: > 0 })
            {
                foreach (var rangeNode in rangesArray)
                {
                    var rangeStr = rangeNode?.GetValue<string>();
                    if (!string.IsNullOrEmpty(rangeStr))
                    {
                        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, rangeStr);
                        cellRange.ApplyStyle(style, styleFlag);
                    }
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

            workbook.Save(outputPath);
            return $"Cells formatted in sheet {sheetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets format information for a cell or range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell or range</param>
    /// <returns>JSON string with cell format details</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither cell nor range is provided, or the range format is invalid
    /// </exception>
    private Task<string> GetCellFormatAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");
            var range = ArgumentHelper.GetStringNullable(arguments, "range");
            var fieldsParam = ArgumentHelper.GetStringNullable(arguments, "fields");

            if (string.IsNullOrEmpty(cell) && string.IsNullOrEmpty(range))
                throw new ArgumentException("Either cell or range is required for get_format operation");

            var cellOrRange = cell ?? range!;

            // Parse fields parameter
            var requestedFields = ParseFields(fieldsParam);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cells = worksheet.Cells;

            try
            {
                var cellRange = ExcelHelper.CreateRange(cells, cellOrRange);
                var startRow = cellRange.FirstRow;
                var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                var startCol = cellRange.FirstColumn;
                var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

                var cellList = new List<Dictionary<string, object?>>();
                for (var row = startRow; row <= endRow; row++)
                for (var col = startCol; col <= endCol; col++)
                {
                    var cellObj = cells[row, col];
                    var style = cellObj.GetStyle();

                    var cellData = new Dictionary<string, object?>
                    {
                        ["cell"] = CellsHelper.CellIndexToName(row, col)
                    };

                    // Value fields (always include cell address)
                    if (requestedFields.Contains("value") || requestedFields.Contains("all"))
                    {
                        cellData["value"] = cellObj.Value?.ToString() ?? "(empty)";
                        cellData["formula"] = cellObj.Formula;
                        cellData["dataType"] = cellObj.Type.ToString();
                    }

                    var formatData = new Dictionary<string, object?>();

                    // Font fields
                    if (requestedFields.Contains("font") || requestedFields.Contains("all"))
                    {
                        formatData["fontName"] = style.Font.Name;
                        formatData["fontSize"] = style.Font.Size;
                        formatData["bold"] = style.Font.IsBold;
                        formatData["italic"] = style.Font.IsItalic;
                        formatData["underline"] = style.Font.Underline.ToString();
                        formatData["strikethrough"] = style.Font.IsStrikeout;
                    }

                    // Color fields
                    if (requestedFields.Contains("color") || requestedFields.Contains("all"))
                    {
                        formatData["fontColor"] = style.Font.Color.ToString();
                        formatData["foregroundColor"] = style.ForegroundColor.ToString();
                        formatData["backgroundColor"] = style.BackgroundColor.ToString();
                        formatData["patternType"] = style.Pattern.ToString();
                    }

                    // Alignment fields
                    if (requestedFields.Contains("alignment") || requestedFields.Contains("all"))
                    {
                        formatData["horizontalAlignment"] = style.HorizontalAlignment.ToString();
                        formatData["verticalAlignment"] = style.VerticalAlignment.ToString();
                    }

                    // Number format fields
                    if (requestedFields.Contains("number") || requestedFields.Contains("all"))
                    {
                        formatData["numberFormat"] = style.Number;
                        formatData["customFormat"] = style.Custom;
                    }

                    // Border fields
                    if (requestedFields.Contains("border") || requestedFields.Contains("all"))
                    {
                        var topBorder = style.Borders[BorderType.TopBorder];
                        var bottomBorder = style.Borders[BorderType.BottomBorder];
                        var leftBorder = style.Borders[BorderType.LeftBorder];
                        var rightBorder = style.Borders[BorderType.RightBorder];

                        formatData["borders"] = new
                        {
                            top = new
                            {
                                lineStyle = topBorder.LineStyle.ToString(), color = topBorder.Color.ToString()
                            },
                            bottom = new
                            {
                                lineStyle = bottomBorder.LineStyle.ToString(), color = bottomBorder.Color.ToString()
                            },
                            left = new
                            {
                                lineStyle = leftBorder.LineStyle.ToString(), color = leftBorder.Color.ToString()
                            },
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
        });
    }

    /// <summary>
    ///     Parses the fields parameter into a set of field names
    /// </summary>
    /// <param name="fieldsParam">Comma-separated list of field names</param>
    /// <returns>HashSet of normalized field names</returns>
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
    ///     Copies column widths and row heights from source sheet to target sheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sourceSheetIndex and targetSheetIndex</param>
    /// <returns>Success message with sheet indices</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range</exception>
    private Task<string> CopySheetFormatAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sourceSheetIndex = ArgumentHelper.GetInt(arguments, "sourceSheetIndex");
            var targetSheetIndex = ArgumentHelper.GetInt(arguments, "targetSheetIndex");
            var copyColumnWidths = ArgumentHelper.GetBool(arguments, "copyColumnWidths", true);
            var copyRowHeights = ArgumentHelper.GetBool(arguments, "copyRowHeights", true);

            using var workbook = new Workbook(path);
            var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
            var targetSheet = ExcelHelper.GetWorksheet(workbook, targetSheetIndex);

            if (copyColumnWidths)
                for (var i = 0; i <= sourceSheet.Cells.MaxDataColumn; i++)
                    targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));

            if (copyRowHeights)
                for (var i = 0; i <= sourceSheet.Cells.MaxDataRow; i++)
                    targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));

            workbook.Save(outputPath);
            return
                $"Sheet format copied from sheet {sourceSheetIndex} to sheet {targetSheetIndex}. Output: {outputPath}";
        });
    }
}