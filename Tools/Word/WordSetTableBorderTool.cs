using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordSetTableBorderTool : IAsposeTool
{
    public string Description => "Set borders for a table or specific cells in a Word document";

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
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            // Cell range (optional, if not specified, applies to entire table)
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based). If specified, only applies to this row"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based). If specified with rowIndex, only applies to this cell"
            },
            // Border settings for each side
            borderTop = new
            {
                type = "boolean",
                description = "Show top border (default: false)"
            },
            borderBottom = new
            {
                type = "boolean",
                description = "Show bottom border (default: false)"
            },
            borderLeft = new
            {
                type = "boolean",
                description = "Show left border (default: false)"
            },
            borderRight = new
            {
                type = "boolean",
                description = "Show right border (default: false)"
            },
            // Border style (applies to all sides)
            lineStyle = new
            {
                type = "string",
                description = "Border line style: none, single, double, dotted, dashed, thick",
                @enum = new[] { "none", "single", "double", "dotted", "dashed", "thick" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Border line width in points (default: 0.5)"
            },
            lineColor = new
            {
                type = "string",
                description = "Border line color (hex format, e.g., '000000' for black, default: black)"
            }
        },
        required = new[] { "path", "tableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var rowIndex = arguments?["rowIndex"]?.GetValue<int?>();
        var columnIndex = arguments?["columnIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        
        if (tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range (total tables: {tables.Count})");
        
        var table = tables[tableIndex];
        
        // Default values
        var defaultLineStyle = arguments?["lineStyle"]?.GetValue<string>() ?? "single";
        var defaultLineWidth = arguments?["lineWidth"]?.GetValue<double>() ?? 0.5;
        var defaultLineColor = arguments?["lineColor"]?.GetValue<string>() ?? "000000";
        
        var lineStyle = GetLineStyle(defaultLineStyle);
        var lineWidth = defaultLineWidth;
        var lineColor = ParseColor(defaultLineColor);
        
        // Determine target cells
        List<Cell> targetCells = new List<Cell>();
        
        if (rowIndex.HasValue && columnIndex.HasValue)
        {
            // Single cell
            if (rowIndex.Value < table.Rows.Count && columnIndex.Value < table.Rows[rowIndex.Value].Cells.Count)
            {
                targetCells.Add(table.Rows[rowIndex.Value].Cells[columnIndex.Value]);
            }
            else
            {
                throw new ArgumentException($"Row {rowIndex.Value} or column {columnIndex.Value} out of range");
            }
        }
        else if (rowIndex.HasValue)
        {
            // Entire row
            if (rowIndex.Value < table.Rows.Count)
            {
                targetCells.AddRange(table.Rows[rowIndex.Value].Cells.Cast<Cell>());
            }
            else
            {
                throw new ArgumentException($"Row {rowIndex.Value} out of range");
            }
        }
        else if (columnIndex.HasValue)
        {
            // Entire column
            foreach (Row row in table.Rows)
            {
                if (columnIndex.Value < row.Cells.Count)
                {
                    targetCells.Add(row.Cells[columnIndex.Value]);
                }
            }
        }
        else
        {
            // Entire table
            foreach (Row row in table.Rows)
            {
                targetCells.AddRange(row.Cells.Cast<Cell>());
            }
        }
        
        // Apply borders to target cells
        foreach (var cell in targetCells)
        {
            var borders = cell.CellFormat.Borders;
            
            // Top border
            if (arguments?["borderTop"]?.GetValue<bool>() == true)
            {
                borders.Top.LineStyle = lineStyle;
                borders.Top.LineWidth = lineWidth;
                borders.Top.Color = lineColor;
            }
            else if (arguments?["borderTop"] != null)
            {
                borders.Top.LineStyle = LineStyle.None;
            }
            
            // Bottom border
            if (arguments?["borderBottom"]?.GetValue<bool>() == true)
            {
                borders.Bottom.LineStyle = lineStyle;
                borders.Bottom.LineWidth = lineWidth;
                borders.Bottom.Color = lineColor;
            }
            else if (arguments?["borderBottom"] != null)
            {
                borders.Bottom.LineStyle = LineStyle.None;
            }
            
            // Left border
            if (arguments?["borderLeft"]?.GetValue<bool>() == true)
            {
                borders.Left.LineStyle = lineStyle;
                borders.Left.LineWidth = lineWidth;
                borders.Left.Color = lineColor;
            }
            else if (arguments?["borderLeft"] != null)
            {
                borders.Left.LineStyle = LineStyle.None;
            }
            
            // Right border
            if (arguments?["borderRight"]?.GetValue<bool>() == true)
            {
                borders.Right.LineStyle = lineStyle;
                borders.Right.LineWidth = lineWidth;
                borders.Right.Color = lineColor;
            }
            else if (arguments?["borderRight"] != null)
            {
                borders.Right.LineStyle = LineStyle.None;
            }
            
            // Note: InsideHorizontal and InsideVertical are not directly available in CellFormat.Borders
            // They are typically handled by setting borders on adjacent cells
            // For now, we skip these settings
        }
        
        doc.Save(outputPath);
        
        var targetDesc = rowIndex.HasValue && columnIndex.HasValue 
            ? $"單元格 ({rowIndex.Value}, {columnIndex.Value})"
            : rowIndex.HasValue 
                ? $"第 {rowIndex.Value} 行"
                : columnIndex.HasValue 
                    ? $"第 {columnIndex.Value} 列"
                    : "整表";
        
        var enabledBorders = new List<string>();
        if (arguments?["borderTop"]?.GetValue<bool>() == true) enabledBorders.Add("上");
        if (arguments?["borderBottom"]?.GetValue<bool>() == true) enabledBorders.Add("下");
        if (arguments?["borderLeft"]?.GetValue<bool>() == true) enabledBorders.Add("左");
        if (arguments?["borderRight"]?.GetValue<bool>() == true) enabledBorders.Add("右");
        
        var bordersDesc = enabledBorders.Count > 0 ? string.Join("、", enabledBorders) : "無";
        
        return await Task.FromResult($"成功設定表格 {tableIndex} {targetDesc} 的邊框：{bordersDesc}");
    }
    
    private LineStyle GetLineStyle(string style)
    {
        return style.ToLower() switch
        {
            "none" => LineStyle.None,
            "single" => LineStyle.Single,
            "double" => LineStyle.Double,
            "dotted" => LineStyle.Dot,
            "dashed" => LineStyle.Single, // Dash not available, use Single instead
            "thick" => LineStyle.Thick,
            _ => LineStyle.Single
        };
    }
    
    private System.Drawing.Color ParseColor(string colorStr)
    {
        if (string.IsNullOrEmpty(colorStr))
            return System.Drawing.Color.Black;
        
        // Remove # if present
        colorStr = colorStr.TrimStart('#');
        
        if (colorStr.Length == 6)
        {
            // RGB hex format
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        
        return System.Drawing.Color.Black;
    }
}

