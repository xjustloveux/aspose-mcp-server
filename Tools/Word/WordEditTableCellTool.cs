using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordEditTableCellTool : IAsposeTool
{
    public string Description => "Edit formatting of a specific table cell (cell-level formatting)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based)"
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based)"
            },
            colIndex = new
            {
                type = "number",
                description = "Column index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            // Cell background
            backgroundColor = new
            {
                type = "string",
                description = "Cell background color (hex format like 'FF0000' for red, or name like 'Red', 'Blue')"
            },
            // Cell alignment
            alignment = new
            {
                type = "string",
                description = "Horizontal alignment: left, center, right",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment: top, center, bottom",
                @enum = new[] { "top", "center", "bottom" }
            },
            // Cell padding
            paddingTop = new
            {
                type = "number",
                description = "Top padding in points"
            },
            paddingBottom = new
            {
                type = "number",
                description = "Bottom padding in points"
            },
            paddingLeft = new
            {
                type = "number",
                description = "Left padding in points"
            },
            paddingRight = new
            {
                type = "number",
                description = "Right padding in points"
            },
            // Text formatting (applies to all runs in the cell)
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '微軟雅黑', 'Arial'). If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text"
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like 'FF0000' for red, or name like 'Red', 'Blue')"
            }
        },
        required = new[] { "path", "tableIndex", "rowIndex", "colIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var colIndex = arguments?["colIndex"]?.GetValue<int>() ?? throw new ArgumentException("colIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"節索引 {sectionIndex} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
        }
        
        var section = doc.Sections[sectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        
        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new ArgumentException($"表格索引 {tableIndex} 超出範圍 (文檔共有 {tables.Count} 個表格)");
        }
        
        var table = tables[tableIndex];
        
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
        {
            throw new ArgumentException($"行索引 {rowIndex} 超出範圍 (表格共有 {table.Rows.Count} 行)");
        }
        
        var row = table.Rows[rowIndex];
        
        if (colIndex < 0 || colIndex >= row.Cells.Count)
        {
            throw new ArgumentException($"列索引 {colIndex} 超出範圍 (該行共有 {row.Cells.Count} 列)");
        }
        
        var cell = row.Cells[colIndex];
        var cellFormat = cell.CellFormat;
        var changes = new List<string>();
        
        // Apply background color
        if (arguments?["backgroundColor"] != null)
        {
            var backgroundColor = arguments["backgroundColor"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                try
                {
                    var bgColor = ParseColor(backgroundColor);
                    cellFormat.Shading.BackgroundPatternColor = bgColor;
                    changes.Add($"背景色：{backgroundColor}");
                }
                catch
                {
                    // Ignore color parsing errors
                }
            }
        }
        
        // Apply horizontal alignment (paragraph alignment)
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            foreach (var para in paragraphs)
            {
                para.ParagraphFormat.Alignment = GetParagraphAlignment(alignment);
            }
            changes.Add($"水平對齊：{alignment}");
        }
        
        // Apply vertical alignment
        if (arguments?["verticalAlignment"] != null)
        {
            var verticalAlignment = arguments["verticalAlignment"]?.GetValue<string>() ?? "top";
            cellFormat.VerticalAlignment = GetCellVerticalAlignment(verticalAlignment);
            changes.Add($"垂直對齊：{verticalAlignment}");
        }
        
        // Apply cell padding
        if (arguments?["paddingTop"] != null)
        {
            var paddingTop = arguments["paddingTop"]?.GetValue<double>();
            if (paddingTop.HasValue)
            {
                cellFormat.TopPadding = paddingTop.Value;
                changes.Add($"上內距：{paddingTop.Value} pt");
            }
        }
        
        if (arguments?["paddingBottom"] != null)
        {
            var paddingBottom = arguments["paddingBottom"]?.GetValue<double>();
            if (paddingBottom.HasValue)
            {
                cellFormat.BottomPadding = paddingBottom.Value;
                changes.Add($"下內距：{paddingBottom.Value} pt");
            }
        }
        
        if (arguments?["paddingLeft"] != null)
        {
            var paddingLeft = arguments["paddingLeft"]?.GetValue<double>();
            if (paddingLeft.HasValue)
            {
                cellFormat.LeftPadding = paddingLeft.Value;
                changes.Add($"左內距：{paddingLeft.Value} pt");
            }
        }
        
        if (arguments?["paddingRight"] != null)
        {
            var paddingRight = arguments["paddingRight"]?.GetValue<double>();
            if (paddingRight.HasValue)
            {
                cellFormat.RightPadding = paddingRight.Value;
                changes.Add($"右內距：{paddingRight.Value} pt");
            }
        }
        
        // Apply text formatting to all runs in the cell
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();
        
        bool hasTextFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) || 
                                 !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue || 
                                 bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);
        
        if (hasTextFormatting)
        {
            // Get all runs in the cell
            var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            
            foreach (var run in runs)
            {
                // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
                if (!string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontNameAscii;
                
                if (!string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontNameFarEast;
                
                if (!string.IsNullOrEmpty(fontName))
                {
                    // If fontNameAscii/FarEast are not set, use fontName for both
                    if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                    {
                        run.Font.Name = fontName;
                    }
                    else
                    {
                        // If only one is set, use fontName as fallback for the other
                        if (string.IsNullOrEmpty(fontNameAscii))
                            run.Font.NameAscii = fontName;
                        if (string.IsNullOrEmpty(fontNameFarEast))
                            run.Font.NameFarEast = fontName;
                    }
                }
                
                if (fontSize.HasValue)
                    run.Font.Size = fontSize.Value;
                
                if (bold.HasValue)
                    run.Font.Bold = bold.Value;
                
                if (italic.HasValue)
                    run.Font.Italic = italic.Value;
                
                if (!string.IsNullOrEmpty(color))
                {
                    try
                    {
                        run.Font.Color = ParseColor(color);
                    }
                    catch
                    {
                        // Ignore color parsing errors
                    }
                }
            }
            
            var textFormatting = new List<string>();
            if (!string.IsNullOrEmpty(fontNameAscii)) textFormatting.Add($"字體（英文）：{fontNameAscii}");
            if (!string.IsNullOrEmpty(fontNameFarEast)) textFormatting.Add($"字體（中文）：{fontNameFarEast}");
            if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast)) 
                textFormatting.Add($"字體：{fontName}");
            if (fontSize.HasValue) textFormatting.Add($"字號：{fontSize.Value} pt");
            if (bold.HasValue && bold.Value) textFormatting.Add("粗體");
            if (italic.HasValue && italic.Value) textFormatting.Add("斜體");
            if (!string.IsNullOrEmpty(color)) textFormatting.Add($"顏色：{color}");
            
            if (textFormatting.Count > 0)
            {
                changes.Add($"文字格式：{string.Join("、", textFormatting)}");
            }
        }
        
        
        doc.Save(outputPath);
        
        var changesDesc = changes.Count > 0 ? string.Join("、", changes) : "格式";
        
        return await Task.FromResult($"成功編輯表格 {tableIndex} 的單元格 [{rowIndex}, {colIndex}] 的{changesDesc}");
    }
    
    private System.Drawing.Color ParseColor(string color)
    {
        // Try to parse as hex color
        if (color.StartsWith("#"))
            color = color.Substring(1);
        
        if (color.Length == 6)
        {
            int r = Convert.ToInt32(color.Substring(0, 2), 16);
            int g = Convert.ToInt32(color.Substring(2, 2), 16);
            int b = Convert.ToInt32(color.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        else
        {
            // Try to parse as color name
            return System.Drawing.Color.FromName(color);
        }
    }
    
    private ParagraphAlignment GetParagraphAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }
    
    private CellVerticalAlignment GetCellVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => CellVerticalAlignment.Top,
            "center" => CellVerticalAlignment.Center,
            "bottom" => CellVerticalAlignment.Bottom,
            _ => CellVerticalAlignment.Top
        };
    }
    
}

