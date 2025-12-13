using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel styles (create, apply, get, format cells)
/// Merges: ExcelCreateStyleTool, ExcelApplyStyleTool, ExcelGetStylesTool, ExcelFormatCellsTool, 
/// ExcelBatchFormatCellsTool, ExcelGetCellFormatTool, ExcelCopySheetFormatTool
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
                description = "Cell range (e.g., 'A1:C5', required for format/get_format)"
            },
            cell = new
            {
                type = "string",
                description = "Cell address or range (required for get_format)"
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
            "format" => await FormatCellsAsync(arguments, path, sheetIndex),
            "get_format" => await GetCellFormatAsync(arguments, path, sheetIndex),
            "copy_sheet_format" => await CopySheetFormatAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> FormatCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();
        var rangesArray = arguments?["ranges"]?.AsArray();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<int?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var fontColor = arguments?["fontColor"]?.GetValue<string>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();
        var numberFormat = arguments?["numberFormat"]?.GetValue<string>();
        var borderStyle = arguments?["borderStyle"]?.GetValue<string>();
        var borderColor = arguments?["borderColor"]?.GetValue<string>();
        var horizontalAlignment = arguments?["horizontalAlignment"]?.GetValue<string>();
        var verticalAlignment = arguments?["verticalAlignment"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var style = workbook.CreateStyle();

        if (!string.IsNullOrEmpty(fontName)) style.Font.Name = fontName;
        if (fontSize.HasValue) style.Font.Size = fontSize.Value;
        if (bold.HasValue) style.Font.IsBold = bold.Value;
        if (italic.HasValue) style.Font.IsItalic = italic.Value;
        if (!string.IsNullOrWhiteSpace(fontColor))
        {
            try
            {
                style.Font.Color = System.Drawing.ColorTranslator.FromHtml(fontColor);
            }
            catch { }
        }
        if (!string.IsNullOrWhiteSpace(backgroundColor))
        {
            try
            {
                style.ForegroundColor = System.Drawing.ColorTranslator.FromHtml(backgroundColor);
                style.Pattern = BackgroundType.Solid;
            }
            catch { }
        }
        if (!string.IsNullOrEmpty(numberFormat))
        {
            // Try to parse as built-in format number, otherwise use Custom
            if (int.TryParse(numberFormat, out int formatNumber))
            {
                style.Number = formatNumber;
            }
            else
            {
                style.Custom = numberFormat;
            }
        }
        if (!string.IsNullOrEmpty(horizontalAlignment))
        {
            style.HorizontalAlignment = horizontalAlignment.ToLower() switch
            {
                "left" => TextAlignmentType.Left,
                "center" => TextAlignmentType.Center,
                "right" => TextAlignmentType.Right,
                _ => TextAlignmentType.Left
            };
        }
        if (!string.IsNullOrEmpty(verticalAlignment))
        {
            style.VerticalAlignment = verticalAlignment.ToLower() switch
            {
                "top" => TextAlignmentType.Top,
                "center" => TextAlignmentType.Center,
                "bottom" => TextAlignmentType.Bottom,
                _ => TextAlignmentType.Center
            };
        }
        
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
            
            System.Drawing.Color borderColorValue = System.Drawing.Color.Black;
            if (!string.IsNullOrWhiteSpace(borderColor))
            {
                try
                {
                    borderColorValue = System.Drawing.ColorTranslator.FromHtml(borderColor);
                }
                catch { }
            }
            
            // Set borders for all sides
            style.SetBorder(BorderType.TopBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.BottomBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.LeftBorder, borderType, borderColorValue);
            style.SetBorder(BorderType.RightBorder, borderType, borderColorValue);
        }

        // Create StyleFlag to specify which style properties to apply
        var styleFlag = new StyleFlag();
        styleFlag.All = true;
        if (!string.IsNullOrEmpty(borderStyle))
        {
            styleFlag.Borders = true;
        }

        if (rangesArray != null && rangesArray.Count > 0)
        {
            foreach (var rangeNode in rangesArray)
            {
                var rangeStr = rangeNode?.GetValue<string>();
                if (!string.IsNullOrEmpty(rangeStr))
                {
                    var cellRange = worksheet.Cells.CreateRange(rangeStr);
                    cellRange.ApplyStyle(style, styleFlag);
                }
            }
        }
        else if (!string.IsNullOrEmpty(range))
        {
            var cellRange = worksheet.Cells.CreateRange(range);
            cellRange.ApplyStyle(style, styleFlag);
        }
        else
        {
            throw new ArgumentException("Either range or ranges must be provided for format operation");
        }

        workbook.Save(path);
        return await Task.FromResult($"Cells formatted in sheet {sheetIndex}: {path}");
    }

    private async Task<string> GetCellFormatAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for get_format operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的單元格格式資訊 ===\n");

        try
        {
            var cellRange = cells.CreateRange(cell);
            var startRow = cellRange.FirstRow;
            var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
            var startCol = cellRange.FirstColumn;
            var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellObj = cells[row, col];
                    var style = cellObj.GetStyle();
                    
                    result.AppendLine($"【單元格 {CellsHelper.CellIndexToName(row, col)}】");
                    result.AppendLine($"值: {cellObj.Value ?? "(空)"}");
                    result.AppendLine($"公式: {cellObj.Formula ?? "(無)"}");
                    result.AppendLine($"數據類型: {cellObj.Type}");
                    result.AppendLine();
                    
                    result.AppendLine("格式資訊:");
                    result.AppendLine($"  字型: {style.Font.Name}, 大小: {style.Font.Size}");
                    result.AppendLine($"  粗體: {style.Font.IsBold}, 斜體: {style.Font.IsItalic}");
                    result.AppendLine($"  底線: {style.Font.Underline}, 刪除線: {style.Font.IsStrikeout}");
                    result.AppendLine($"  字型顏色: {style.Font.Color}");
                    result.AppendLine($"  背景色: {style.BackgroundColor}");
                    result.AppendLine($"  數字格式: {style.Number}");
                    result.AppendLine($"  水平對齊: {style.HorizontalAlignment}");
                    result.AppendLine($"  垂直對齊: {style.VerticalAlignment}");
                    
                    // Add border information
                    result.AppendLine("  邊框資訊:");
                    var topBorder = style.Borders[BorderType.TopBorder];
                    var bottomBorder = style.Borders[BorderType.BottomBorder];
                    var leftBorder = style.Borders[BorderType.LeftBorder];
                    var rightBorder = style.Borders[BorderType.RightBorder];
                    
                    result.AppendLine($"    上邊框: {topBorder.LineStyle} ({topBorder.Color})");
                    result.AppendLine($"    下邊框: {bottomBorder.LineStyle} ({bottomBorder.Color})");
                    result.AppendLine($"    左邊框: {leftBorder.LineStyle} ({leftBorder.Color})");
                    result.AppendLine($"    右邊框: {rightBorder.LineStyle} ({rightBorder.Color})");
                    result.AppendLine();
                }
            }
        }
        catch
        {
            throw new ArgumentException($"無效的單元格範圍: {cell}");
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> CopySheetFormatAsync(JsonObject? arguments, string path)
    {
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceSheetIndex is required for copy_sheet_format operation");
        var targetSheetIndex = arguments?["targetSheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetSheetIndex is required for copy_sheet_format operation");
        var copyColumnWidths = arguments?["copyColumnWidths"]?.GetValue<bool?>() ?? true;
        var copyRowHeights = arguments?["copyRowHeights"]?.GetValue<bool?>() ?? true;

        using var workbook = new Workbook(path);
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
        var targetSheet = ExcelHelper.GetWorksheet(workbook, targetSheetIndex);

        if (copyColumnWidths)
        {
            for (int i = 0; i <= sourceSheet.Cells.MaxDataColumn; i++)
            {
                targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));
            }
        }

        if (copyRowHeights)
        {
            for (int i = 0; i <= sourceSheet.Cells.MaxDataRow; i++)
            {
                targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));
            }
        }

        workbook.Save(path);
        return await Task.FromResult($"Sheet format copied from sheet {sourceSheetIndex} to sheet {targetSheetIndex}: {path}");
    }
}

