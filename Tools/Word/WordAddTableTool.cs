using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordAddTableTool : IAsposeTool
{
    public string Description => "Add table to Word document (supports formatting, colors, merge cells, vertical alignment, Chinese/English fonts)";

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
            rows = new
            {
                type = "number",
                description = "Number of rows"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns"
            },
            data = new
            {
                type = "array",
                description = "Table data (array of arrays)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            headerRow = new
            {
                type = "boolean",
                description = "First row is header (default: false)"
            },
            headerBackgroundColor = new
            {
                type = "string",
                description = "Header row background color (hex format)"
            },
            rowBackgroundColors = new
            {
                type = "object",
                description = "Background colors for specific rows, e.g., {\"0\": \"dbe5f1\", \"1\": \"ffffff\"}"
            },
            columnBackgroundColors = new
            {
                type = "object",
                description = "Background colors for specific columns, e.g., {\"0\": \"f2f2f2\"}"
            },
            cellBackgroundColors = new
            {
                type = "array",
                description = "Background colors for specific cells [row, col, color]",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            rowStyles = new
            {
                type = "object",
                description = "Text styles for specific rows, e.g., {\"0\": {\"bold\": true, \"color\": \"ffffff\"}}"
            },
            mergeCells = new
            {
                type = "array",
                description = "Cells to merge, e.g., [{\"startRow\": 0, \"endRow\": 1, \"startCol\": 0, \"endCol\": 0}]",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        startRow = new { type = "number" },
                        endRow = new { type = "number" },
                        startCol = new { type = "number" },
                        endCol = new { type = "number" }
                    }
                }
            },
            borderStyle = new
            {
                type = "string",
                description = "Border style: none, single, double, dotted (default: single)",
                @enum = new[] { "none", "single", "double", "dotted" }
            },
            alignment = new
            {
                type = "string",
                description = "Table horizontal alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Cell vertical alignment: top, center, bottom (default: center)",
                @enum = new[] { "top", "center", "bottom" }
            },
            cellPadding = new
            {
                type = "number",
                description = "Cell padding in points (default: 5)"
            },
            tableFontName = new
            {
                type = "string",
                description = "Font name for all table cells (e.g., '標楷體', 'Arial'). Sets both ASCII and Far East fonts if not specified separately."
            },
            tableFontSize = new
            {
                type = "number",
                description = "Font size for all table cells in points (e.g., 12)"
            },
            tableFontNameAscii = new
            {
                type = "string",
                description = "Font name for English text in table cells (e.g., 'Times New Roman')"
            },
            tableFontNameFarEast = new
            {
                type = "string",
                description = "Font name for Chinese/Japanese/Korean text in table cells (e.g., '標楷體')"
            },
            headerFontName = new
            {
                type = "string",
                description = "Font name for header row (optional, defaults to tableFontName)"
            },
            headerFontSize = new
            {
                type = "number",
                description = "Font size for header row in points (optional, defaults to tableFontSize)"
            },
            headerFontNameAscii = new
            {
                type = "string",
                description = "Font name for English text in header row (optional, defaults to tableFontNameAscii)"
            },
            headerFontNameFarEast = new
            {
                type = "string",
                description = "Font name for Chinese/Japanese/Korean text in header row (optional, defaults to tableFontNameFarEast)"
            }
        },
        required = new[] { "path", "rows", "columns" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var rows = arguments?["rows"]?.GetValue<int>() ?? throw new ArgumentException("rows is required");
        var columns = arguments?["columns"]?.GetValue<int>() ?? throw new ArgumentException("columns is required");
        var headerRow = arguments?["headerRow"]?.GetValue<bool>() ?? false;
        var headerBgColor = arguments?["headerBackgroundColor"]?.GetValue<string>();
        var borderStyle = arguments?["borderStyle"]?.GetValue<string>() ?? "single";
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";
        var verticalAlignment = arguments?["verticalAlignment"]?.GetValue<string>() ?? "center";
        var cellPadding = arguments?["cellPadding"]?.GetValue<double>() ?? 5.0;
        var tableFontName = arguments?["tableFontName"]?.GetValue<string>();
        var tableFontSize = arguments?["tableFontSize"]?.GetValue<double?>();
        var tableFontNameAscii = arguments?["tableFontNameAscii"]?.GetValue<string>();
        var tableFontNameFarEast = arguments?["tableFontNameFarEast"]?.GetValue<string>();
        var headerFontName = arguments?["headerFontName"]?.GetValue<string>() ?? tableFontName;
        var headerFontSize = arguments?["headerFontSize"]?.GetValue<double?>() ?? tableFontSize;
        var headerFontNameAscii = arguments?["headerFontNameAscii"]?.GetValue<string>() ?? tableFontNameAscii;
        var headerFontNameFarEast = arguments?["headerFontNameFarEast"]?.GetValue<string>() ?? tableFontNameFarEast;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();
        
        // Insert a new paragraph before the table and reset its indentation
        // This ensures the table is not affected by any surrounding paragraph's indentation
        builder.InsertParagraph();
        builder.CurrentParagraph.ParagraphFormat.LeftIndent = 0;
        builder.CurrentParagraph.ParagraphFormat.RightIndent = 0;
        builder.CurrentParagraph.ParagraphFormat.FirstLineIndent = 0;

        // Parse data
        string[][]? data = null;
        if (arguments?.ContainsKey("data") == true)
        {
            try
            {
                var dataJson = arguments["data"]?.ToJsonString();
                if (!string.IsNullOrEmpty(dataJson))
                {
                    data = JsonSerializer.Deserialize<string[][]>(dataJson);
                }
            }
            catch { }
        }

        // Parse row background colors
        var rowBgColors = ParseColorDictionary(arguments?["rowBackgroundColors"]);
        
        // Parse column background colors
        var columnBgColors = ParseColorDictionary(arguments?["columnBackgroundColors"]);

        // Parse cell background colors
        var cellColors = ParseCellColors(arguments?["cellBackgroundColors"]);

        // Parse row styles
        var rowStyles = ParseRowStyles(arguments?["rowStyles"]);

        // Parse merge cells
        var mergeCells = ParseMergeCells(arguments?["mergeCells"]);

        // Create table
        var table = builder.StartTable();
        var cells = new Dictionary<(int row, int col), Cell>();

        // Build table
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                var cell = builder.InsertCell();
                cells[(i, j)] = cell;

                // Set cell content
                string cellText = "";
                if (data != null && i < data.Length && j < data[i].Length)
                {
                    cellText = data[i][j];
                }
                else
                {
                    cellText = $"Cell {i + 1},{j + 1}";
                }

                // Apply table font settings first
                bool isHeaderRow = headerRow && i == 0;
                
                // Apply font name (unified or separate ASCII/FarEast)
                if (isHeaderRow)
                {
                    if (!string.IsNullOrEmpty(headerFontName))
                        builder.Font.Name = headerFontName;
                    if (!string.IsNullOrEmpty(headerFontNameAscii))
                        builder.Font.NameAscii = headerFontNameAscii;
                    if (!string.IsNullOrEmpty(headerFontNameFarEast))
                        builder.Font.NameFarEast = headerFontNameFarEast;
                }
                else
                {
                    if (!string.IsNullOrEmpty(tableFontName))
                        builder.Font.Name = tableFontName;
                    if (!string.IsNullOrEmpty(tableFontNameAscii))
                        builder.Font.NameAscii = tableFontNameAscii;
                    if (!string.IsNullOrEmpty(tableFontNameFarEast))
                        builder.Font.NameFarEast = tableFontNameFarEast;
                }
                
                if (isHeaderRow && headerFontSize.HasValue)
                    builder.Font.Size = headerFontSize.Value;
                else if (tableFontSize.HasValue)
                    builder.Font.Size = tableFontSize.Value;

                // Apply row styles (overrides table font settings if specified)
                if (rowStyles.ContainsKey(i))
                {
                    var style = rowStyles[i];
                    if (style.ContainsKey("bold") && bool.TryParse(style["bold"], out bool bold))
                        builder.Font.Bold = bold;
                    if (style.ContainsKey("italic") && bool.TryParse(style["italic"], out bool italic))
                        builder.Font.Italic = italic;
                    if (style.ContainsKey("color"))
                        builder.Font.Color = ParseColor(style["color"]);
                }

                builder.Write(cellText);

                // Reset font
                builder.Font.Bold = false;
                builder.Font.Italic = false;
                builder.Font.Color = System.Drawing.Color.Black;
                builder.Font.Name = "Calibri";
                builder.Font.Size = 11;

                // Set cell formatting
                cell.CellFormat.SetPaddings(cellPadding, cellPadding, cellPadding, cellPadding);

                // Reset background color
                cell.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Empty;

                // Apply vertical alignment
                cell.CellFormat.VerticalAlignment = GetVerticalAlignment(verticalAlignment);

                // Priority: cell color > column color > row color > header color
                bool hasColor = false;

                // Cell specific color (highest priority)
                var cellColorMatch = cellColors.FirstOrDefault(c => c.row == i && c.col == j);
                if (cellColorMatch != default)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = ParseColor(cellColorMatch.color);
                    hasColor = true;
                }

                // Column color
                if (!hasColor && columnBgColors.ContainsKey(j))
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = ParseColor(columnBgColors[j]);
                    hasColor = true;
                }

                // Row color
                if (!hasColor && rowBgColors.ContainsKey(i))
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = ParseColor(rowBgColors[i]);
                    hasColor = true;
                }

                // Header row color (lowest priority)
                if (!hasColor && headerRow && i == 0 && !string.IsNullOrEmpty(headerBgColor))
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = ParseColor(headerBgColor);
                }

                // Set border style
                if (borderStyle != "none")
                {
                    var lineStyle = borderStyle switch
                    {
                        "double" => LineStyle.Double,
                        "dotted" => LineStyle.Dot,
                        _ => LineStyle.Single
                    };

                    cell.CellFormat.Borders.LineStyle = lineStyle;
                    cell.CellFormat.Borders.Color = System.Drawing.Color.Black;
                }
                else
                {
                    cell.CellFormat.Borders.LineStyle = LineStyle.None;
                }
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Apply cell merges
        foreach (var merge in mergeCells)
        {
            try
            {
                var startCell = cells[(merge.startRow, merge.startCol)];
                var endCell = cells[(merge.endRow, merge.endCol)];
                
                // Horizontal merge
                if (merge.startRow == merge.endRow && merge.startCol != merge.endCol)
                {
                    startCell.CellFormat.HorizontalMerge = CellMerge.First;
                    for (int col = merge.startCol + 1; col <= merge.endCol; col++)
                    {
                        if (cells.ContainsKey((merge.startRow, col)))
                        {
                            cells[(merge.startRow, col)].CellFormat.HorizontalMerge = CellMerge.Previous;
                        }
                    }
                }
                // Vertical merge
                else if (merge.startCol == merge.endCol && merge.startRow != merge.endRow)
                {
                    startCell.CellFormat.VerticalMerge = CellMerge.First;
                    for (int row = merge.startRow + 1; row <= merge.endRow; row++)
                    {
                        if (cells.ContainsKey((row, merge.startCol)))
                        {
                            cells[(row, merge.startCol)].CellFormat.VerticalMerge = CellMerge.Previous;
                        }
                    }
                }
            }
            catch
            {
                // Skip invalid merge
            }
        }

        // Set table alignment
        table.Alignment = alignment.ToLower() switch
        {
            "center" => TableAlignment.Center,
            "right" => TableAlignment.Right,
            _ => TableAlignment.Left
        };

        doc.Save(outputPath);

        var result = $"成功添加增強表格 ({rows} 行 x {columns} 列)\n";
        if (headerRow) result += "標題行: 是\n";
        if (!string.IsNullOrEmpty(tableFontName)) result += $"表格字型: {tableFontName}\n";
        if (tableFontSize.HasValue) result += $"表格字號: {tableFontSize.Value} pt\n";
        if (headerRow && !string.IsNullOrEmpty(headerFontName)) result += $"標題行字型: {headerFontName}\n";
        if (headerRow && headerFontSize.HasValue) result += $"標題行字號: {headerFontSize.Value} pt\n";
        if (rowBgColors.Any()) result += $"行背景色: {rowBgColors.Count} 行\n";
        if (columnBgColors.Any()) result += $"列背景色: {columnBgColors.Count} 列\n";
        if (cellColors.Any()) result += $"儲存格背景色: {cellColors.Count} 個\n";
        if (rowStyles.Any()) result += $"行樣式: {rowStyles.Count} 行\n";
        if (mergeCells.Any()) result += $"合併儲存格: {mergeCells.Count} 個\n";
        result += $"垂直對齊: {verticalAlignment}\n";
        result += $"邊框樣式: {borderStyle}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private Dictionary<int, string> ParseColorDictionary(JsonNode? node)
    {
        var result = new Dictionary<int, string>();
        if (node == null) return result;

        try
        {
            var jsonObj = node.AsObject();
            foreach (var kvp in jsonObj)
            {
                if (int.TryParse(kvp.Key, out int key))
                {
                    result[key] = kvp.Value?.GetValue<string>() ?? "";
                }
            }
        }
        catch { }

        return result;
    }

    private List<(int row, int col, string color)> ParseCellColors(JsonNode? node)
    {
        var result = new List<(int row, int col, string color)>();
        if (node == null) return result;

        try
        {
            var jsonStr = node.ToJsonString();
            var arr = JsonSerializer.Deserialize<JsonElement[][]>(jsonStr);
            if (arr != null)
            {
                foreach (var item in arr)
                {
                    if (item.Length >= 3)
                    {
                        result.Add((item[0].GetInt32(), item[1].GetInt32(), item[2].GetString() ?? ""));
                    }
                }
            }
        }
        catch { }

        return result;
    }

    private Dictionary<int, Dictionary<string, string>> ParseRowStyles(JsonNode? node)
    {
        var result = new Dictionary<int, Dictionary<string, string>>();
        if (node == null) return result;

        try
        {
            var jsonObj = node.AsObject();
            foreach (var kvp in jsonObj)
            {
                if (int.TryParse(kvp.Key, out int rowIdx))
                {
                    var styleDict = new Dictionary<string, string>();
                    var styleObj = kvp.Value?.AsObject();
                    if (styleObj != null)
                    {
                        foreach (var styleProp in styleObj)
                        {
                            styleDict[styleProp.Key] = styleProp.Value?.ToString() ?? "";
                        }
                    }
                    result[rowIdx] = styleDict;
                }
            }
        }
        catch { }

        return result;
    }

    private List<(int startRow, int endRow, int startCol, int endCol)> ParseMergeCells(JsonNode? node)
    {
        var result = new List<(int startRow, int endRow, int startCol, int endCol)>();
        if (node == null) return result;

        try
        {
            var jsonStr = node.ToJsonString();
            var arr = JsonSerializer.Deserialize<JsonElement[]>(jsonStr);
            if (arr != null)
            {
                foreach (var item in arr)
                {
                    if (item.TryGetProperty("startRow", out var sr) &&
                        item.TryGetProperty("endRow", out var er) &&
                        item.TryGetProperty("startCol", out var sc) &&
                        item.TryGetProperty("endCol", out var ec))
                    {
                        result.Add((sr.GetInt32(), er.GetInt32(), sc.GetInt32(), ec.GetInt32()));
                    }
                }
            }
        }
        catch { }

        return result;
    }

    private CellVerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => CellVerticalAlignment.Top,
            "bottom" => CellVerticalAlignment.Bottom,
            _ => CellVerticalAlignment.Center
        };
    }

    private System.Drawing.Color ParseColor(string color)
    {
        try
        {
            if (color.StartsWith("#"))
                color = color.Substring(1);

            if (color.Length == 8)
            {
                int a = Convert.ToInt32(color.Substring(0, 2), 16);
                int r = Convert.ToInt32(color.Substring(2, 2), 16);
                int g = Convert.ToInt32(color.Substring(4, 2), 16);
                int b = Convert.ToInt32(color.Substring(6, 2), 16);
                return System.Drawing.Color.FromArgb(a, r, g, b);
            }
            else if (color.Length == 6)
            {
                int r = Convert.ToInt32(color.Substring(0, 2), 16);
                int g = Convert.ToInt32(color.Substring(2, 2), 16);
                int b = Convert.ToInt32(color.Substring(4, 2), 16);
                return System.Drawing.Color.FromArgb(r, g, b);
            }
            else
            {
                return System.Drawing.Color.FromName(color);
            }
        }
        catch
        {
            return System.Drawing.Color.White;
        }
    }
}

