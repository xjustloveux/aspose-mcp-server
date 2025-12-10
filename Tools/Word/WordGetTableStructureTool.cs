using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordGetTableStructureTool : IAsposeTool
{
    public string Description => "Get detailed structure and formatting of a specific table (useful for copying table structure from one document to another)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, default: 0)"
            },
            includeContent = new
            {
                type = "boolean",
                description = "Include cell content (default: true)"
            },
            includeCellFormatting = new
            {
                type = "boolean",
                description = "Include cell-level formatting details (default: true)"
            }
        },
        required = new[] { "path", "tableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var includeContent = arguments?["includeContent"]?.GetValue<bool>() ?? true;
        var includeCellFormatting = arguments?["includeCellFormatting"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"節索引 {sectionIndex} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
        }
        
        var section = doc.Sections[sectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        
        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new ArgumentException($"表格索引 {tableIndex} 超出範圍 (節 {sectionIndex} 共有 {tables.Count} 個表格)");
        }
        
        var table = tables[tableIndex];
        var result = new StringBuilder();
        
        result.AppendLine($"=== 表格 #{tableIndex} 結構資訊 ===\n");

        // Basic info
        result.AppendLine("【基本資訊】");
        result.AppendLine($"行數: {table.Rows.Count}");
        if (table.Rows.Count > 0)
        {
            result.AppendLine($"列數: {table.Rows[0].Cells.Count}");
        }
        result.AppendLine();

        // Table format
        result.AppendLine("【表格格式】");
        result.AppendLine($"對齊方式: {table.Alignment}");
        result.AppendLine($"表格樣式: {table.Style?.Name ?? "無"}");
        result.AppendLine($"左縮排: {table.LeftIndent:F2} pt");
        
        if (table.PreferredWidth.Type != PreferredWidthType.Auto)
        {
            result.AppendLine($"寬度: {table.PreferredWidth.Value} ({table.PreferredWidth.Type})");
        }
        
        result.AppendLine($"允許自動調整: {table.AllowAutoFit}");
        
        if (table.CellSpacing > 0)
        {
            result.AppendLine($"儲存格間距: {table.CellSpacing:F2} pt");
        }
        result.AppendLine();

        // Borders
        result.AppendLine("【表格邊框】");
        var borders = table.FirstRow.Cells[0].CellFormat.Borders;
        if (borders.Top.LineStyle != LineStyle.None)
            result.AppendLine($"上邊框: {borders.Top.LineStyle}, {borders.Top.LineWidth} pt");
        if (borders.Bottom.LineStyle != LineStyle.None)
            result.AppendLine($"下邊框: {borders.Bottom.LineStyle}, {borders.Bottom.LineWidth} pt");
        if (borders.Left.LineStyle != LineStyle.None)
            result.AppendLine($"左邊框: {borders.Left.LineStyle}, {borders.Left.LineWidth} pt");
        if (borders.Right.LineStyle != LineStyle.None)
            result.AppendLine($"右邊框: {borders.Right.LineStyle}, {borders.Right.LineWidth} pt");
        result.AppendLine();

        // Row structure
        result.AppendLine("【行結構】");
        for (int i = 0; i < table.Rows.Count; i++)
        {
            var row = table.Rows[i];
            result.AppendLine($"\n  行 {i}:");
            result.AppendLine($"    儲存格數: {row.Cells.Count}");
            result.AppendLine($"    行高: {(row.RowFormat.Height > 0 ? $"{row.RowFormat.Height:F2} pt" : "自動")}");
            result.AppendLine($"    行高規則: {row.RowFormat.HeightRule}");
            
            if (row.RowFormat.HeadingFormat)
                result.AppendLine($"    標題行: 是");
            
            if (i == 0)
            {
                result.AppendLine($"\n    儲存格詳細資訊:");
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    var cell = row.Cells[j];
                    result.AppendLine($"\n      儲存格 {j}:");
                    
                    if (cell.CellFormat.Width > 0)
                        result.AppendLine($"        寬度: {cell.CellFormat.Width:F2} pt");
                    
                    result.AppendLine($"        垂直對齊: {cell.CellFormat.VerticalAlignment}");
                    
                    if (cell.CellFormat.HorizontalMerge != CellMerge.None)
                        result.AppendLine($"        水平合併: {cell.CellFormat.HorizontalMerge}");
                    
                    if (cell.CellFormat.VerticalMerge != CellMerge.None)
                        result.AppendLine($"        垂直合併: {cell.CellFormat.VerticalMerge}");
                    
                    var shading = cell.CellFormat.Shading;
                    if (shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.Empty.ToArgb() &&
                        shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.White.ToArgb())
                    {
                        var color = shading.BackgroundPatternColor;
                        result.AppendLine($"        背景色: #{color.R:X2}{color.G:X2}{color.B:X2}");
                    }
                }
            }
        }
        result.AppendLine();

        // Content preview
        if (includeContent)
        {
            result.AppendLine("【內容預覽】");
            for (int i = 0; i < Math.Min(table.Rows.Count, 5); i++)
            {
                var row = table.Rows[i];
                result.Append($"  行 {i}: | ");
                
                for (int j = 0; j < row.Cells.Count; j++)
                {
                    var cell = row.Cells[j];
                    var cellText = cell.GetText().Trim().Replace("\r", "").Replace("\n", " ");
                    
                    // Truncate long text
                    if (cellText.Length > 30)
                        cellText = cellText.Substring(0, 27) + "...";
                    
                    result.Append($"{cellText} | ");
                }
                result.AppendLine();
            }
            
            if (table.Rows.Count > 5)
            {
                result.AppendLine($"  ... 還有 {table.Rows.Count - 5} 行（已省略）");
            }
            result.AppendLine();
        }

        // Cell formatting details
        if (includeCellFormatting && table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0)
        {
            result.AppendLine("【第一個儲存格的詳細格式】");
            var cell = table.Rows[0].Cells[0];
            
            result.AppendLine($"上內距: {cell.CellFormat.TopPadding:F2} pt");
            result.AppendLine($"下內距: {cell.CellFormat.BottomPadding:F2} pt");
            result.AppendLine($"左內距: {cell.CellFormat.LeftPadding:F2} pt");
            result.AppendLine($"右內距: {cell.CellFormat.RightPadding:F2} pt");
            
            // Font info from first paragraph
            var para = cell.FirstParagraph;
            if (para != null && para.Runs.Count > 0)
            {
                var run = para.Runs[0];
                result.AppendLine($"\n字型資訊:");
                
                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    result.AppendLine($"  字體（英文）: {run.Font.NameAscii}");
                    result.AppendLine($"  字體（中文）: {run.Font.NameFarEast}");
                }
                else
                {
                    result.AppendLine($"  字體: {run.Font.Name}");
                }
                
                result.AppendLine($"  字號: {run.Font.Size} pt");
                
                if (run.Font.Bold) result.AppendLine($"  粗體: 是");
                if (run.Font.Italic) result.AppendLine($"  斜體: 是");
                
                result.AppendLine($"  段落對齊: {para.ParagraphFormat.Alignment}");
            }
            result.AppendLine();
        }

        // JSON format for creating similar table
        result.AppendLine("【JSON 格式（可用於 word_add_table_enhanced）】");
        result.AppendLine("{");
        result.AppendLine($"  \"rows\": {table.Rows.Count},");
        result.AppendLine($"  \"columns\": {(table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0)},");
        result.AppendLine($"  \"tableAlignment\": \"{table.Alignment.ToString().ToLower()}\",");
        
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0)
        {
            var cell = table.Rows[0].Cells[0];
            result.AppendLine($"  \"cellPaddingTop\": {cell.CellFormat.TopPadding:F2},");
            result.AppendLine($"  \"cellPaddingBottom\": {cell.CellFormat.BottomPadding:F2},");
            result.AppendLine($"  \"cellPaddingLeft\": {cell.CellFormat.LeftPadding:F2},");
            result.AppendLine($"  \"cellPaddingRight\": {cell.CellFormat.RightPadding:F2}");
        }
        
        result.AppendLine("}");

        return await Task.FromResult(result.ToString());
    }
}

