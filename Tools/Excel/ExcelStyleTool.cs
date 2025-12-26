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
    public string Description => @"Manage Excel styles. Supports 3 operations: format, get_format, copy_sheet_format.

Usage examples:
- Format cells: excel_style(operation='format', path='book.xlsx', range='A1:B10', fontName='Arial', fontSize=12, bold=true)
- Get format: excel_style(operation='get_format', path='book.xlsx', range='A1')
- Copy sheet format: excel_style(operation='copy_sheet_format', path='book.xlsx', sourceSheetIndex=0, targetSheetIndex=1)";

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
                description = "Background color (hex format like '#FFFF00', optional)"
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
    /// <param name="arguments">JSON arguments containing ranges array and various format properties</param>
    /// <returns>Success message with formatted range count</returns>
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

            if (!string.IsNullOrWhiteSpace(backgroundColor))
            {
                // Parse color with error handling - throws ArgumentException on failure
                style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, true);
                style.Pattern = BackgroundType.Solid;
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
    ///     Gets format information for a cell
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <returns>JSON string with cell format details</returns>
    private Task<string> GetCellFormatAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");
            var range = ArgumentHelper.GetStringNullable(arguments, "range");

            if (string.IsNullOrEmpty(cell) && string.IsNullOrEmpty(range))
                throw new ArgumentException("Either cell or range is required for get_format operation");

            var cellOrRange = cell ?? range!;

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

                var cellList = new List<object>();
                for (var row = startRow; row <= endRow; row++)
                for (var col = startCol; col <= endCol; col++)
                {
                    var cellObj = cells[row, col];
                    var style = cellObj.GetStyle();

                    var topBorder = style.Borders[BorderType.TopBorder];
                    var bottomBorder = style.Borders[BorderType.BottomBorder];
                    var leftBorder = style.Borders[BorderType.LeftBorder];
                    var rightBorder = style.Borders[BorderType.RightBorder];

                    cellList.Add(new
                    {
                        cell = CellsHelper.CellIndexToName(row, col),
                        value = cellObj.Value?.ToString() ?? "(empty)",
                        formula = cellObj.Formula,
                        dataType = cellObj.Type.ToString(),
                        format = new
                        {
                            fontName = style.Font.Name,
                            fontSize = style.Font.Size,
                            bold = style.Font.IsBold,
                            italic = style.Font.IsItalic,
                            underline = style.Font.Underline.ToString(),
                            strikethrough = style.Font.IsStrikeout,
                            fontColor = style.Font.Color.ToString(),
                            backgroundColor = style.BackgroundColor.ToString(),
                            numberFormat = style.Number,
                            horizontalAlignment = style.HorizontalAlignment.ToString(),
                            verticalAlignment = style.VerticalAlignment.ToString(),
                            borders = new
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
                            }
                        }
                    });
                }

                var result = new
                {
                    count = cellList.Count,
                    worksheetName = worksheet.Name,
                    range = cellOrRange,
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
    ///     Copies format from source sheet to destination sheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sourceSheetIndex and destSheetIndex</param>
    /// <returns>Success message with sheet names</returns>
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