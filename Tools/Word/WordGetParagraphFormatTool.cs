using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetParagraphFormatTool : IAsposeTool
{
    public string Description => "Get detailed formatting information of a specific paragraph (useful for copying format from one document to another)";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based)"
            },
            includeRunDetails = new
            {
                type = "boolean",
                description = "Include detailed run-level formatting (default: true)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var includeRunDetails = arguments?["includeRunDetails"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"無法找到索引 {paragraphIndex} 的段落");
        }

        var result = new StringBuilder();
        result.AppendLine($"=== 段落 #{paragraphIndex} 格式資訊 ===\n");

        // Basic info
        result.AppendLine("【基本資訊】");
        result.AppendLine($"段落文字: {para.GetText().Trim()}");
        result.AppendLine($"文字長度: {para.GetText().Trim().Length} 字元");
        result.AppendLine($"Run 數量: {para.Runs.Count}");
        result.AppendLine();

        // Paragraph format
        var format = para.ParagraphFormat;
        result.AppendLine("【段落格式】");
        result.AppendLine($"樣式名稱: {format.StyleName}");
        result.AppendLine($"對齊方式: {format.Alignment}");
        result.AppendLine($"左縮排: {format.LeftIndent:F2} pt ({format.LeftIndent / 28.35:F2} cm)");
        result.AppendLine($"右縮排: {format.RightIndent:F2} pt ({format.RightIndent / 28.35:F2} cm)");
        result.AppendLine($"首行縮排: {format.FirstLineIndent:F2} pt ({format.FirstLineIndent / 28.35:F2} cm)");
        result.AppendLine($"段前間距: {format.SpaceBefore:F2} pt");
        result.AppendLine($"段後間距: {format.SpaceAfter:F2} pt");
        result.AppendLine($"行距: {format.LineSpacing:F2} pt");
        result.AppendLine($"行距規則: {format.LineSpacingRule}");
        result.AppendLine();

        // List format
        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine("【列表格式】");
            result.AppendLine($"是列表項: 是");
            result.AppendLine($"列表層級: {para.ListFormat.ListLevelNumber}");
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"列表 ID: {para.ListFormat.List.ListId}");
            }
            result.AppendLine();
        }

        // Borders
        if (format.Borders.Count > 0)
        {
            result.AppendLine("【邊框】");
            if (format.Borders.Top.LineStyle != LineStyle.None)
                result.AppendLine($"上邊框: {format.Borders.Top.LineStyle}, {format.Borders.Top.LineWidth} pt, 顏色: {format.Borders.Top.Color.Name}");
            if (format.Borders.Bottom.LineStyle != LineStyle.None)
                result.AppendLine($"下邊框: {format.Borders.Bottom.LineStyle}, {format.Borders.Bottom.LineWidth} pt, 顏色: {format.Borders.Bottom.Color.Name}");
            if (format.Borders.Left.LineStyle != LineStyle.None)
                result.AppendLine($"左邊框: {format.Borders.Left.LineStyle}, {format.Borders.Left.LineWidth} pt, 顏色: {format.Borders.Left.Color.Name}");
            if (format.Borders.Right.LineStyle != LineStyle.None)
                result.AppendLine($"右邊框: {format.Borders.Right.LineStyle}, {format.Borders.Right.LineWidth} pt, 顏色: {format.Borders.Right.Color.Name}");
            result.AppendLine();
        }

        // Shading
        if (format.Shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.Empty.ToArgb())
        {
            result.AppendLine("【背景色】");
            var color = format.Shading.BackgroundPatternColor;
            result.AppendLine($"背景色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            result.AppendLine();
        }

        // Tab stops
        if (format.TabStops.Count > 0)
        {
            result.AppendLine("【Tab 停駐點】");
            for (int i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                result.AppendLine($"  Tab {i + 1}: 位置={tab.Position:F2} pt, 對齊={tab.Alignment}, 前導字元={tab.Leader}");
            }
            result.AppendLine();
        }

        // Font formatting from first run
        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            result.AppendLine("【字型格式（第一個 Run）】");
            
            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                result.AppendLine($"字體（英文）: {firstRun.Font.NameAscii}");
                result.AppendLine($"字體（中文）: {firstRun.Font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"字體: {firstRun.Font.Name}");
            }
            
            result.AppendLine($"字號: {firstRun.Font.Size} pt");
            
            if (firstRun.Font.Bold) result.AppendLine("粗體: 是");
            if (firstRun.Font.Italic) result.AppendLine("斜體: 是");
            if (firstRun.Font.Underline != Underline.None) result.AppendLine($"底線: {firstRun.Font.Underline}");
            if (firstRun.Font.StrikeThrough) result.AppendLine("刪除線: 是");
            if (firstRun.Font.Superscript) result.AppendLine("上標: 是");
            if (firstRun.Font.Subscript) result.AppendLine("下標: 是");
            
            if (firstRun.Font.Color.ToArgb() != System.Drawing.Color.Empty.ToArgb())
            {
                var color = firstRun.Font.Color;
                result.AppendLine($"顏色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            }
            
            if (firstRun.Font.HighlightColor != System.Drawing.Color.Empty)
            {
                result.AppendLine($"螢光筆: {firstRun.Font.HighlightColor.Name}");
            }
            result.AppendLine();
        }

        // Run details
        if (includeRunDetails && para.Runs.Count > 1)
        {
            result.AppendLine("【Run 詳細資訊】");
            result.AppendLine($"共 {para.Runs.Count} 個 Run:");
            
            for (int i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                result.AppendLine($"\n  Run #{i}:");
                result.AppendLine($"    文字: {run.Text.Replace("\r", "\\r").Replace("\n", "\\n")}");
                
                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    result.AppendLine($"    字體（英文）: {run.Font.NameAscii}");
                    result.AppendLine($"    字體（中文）: {run.Font.NameFarEast}");
                }
                else
                {
                    result.AppendLine($"    字體: {run.Font.Name}");
                }
                
                result.AppendLine($"    字號: {run.Font.Size} pt");
                
                var styles = new List<string>();
                if (run.Font.Bold) styles.Add("粗體");
                if (run.Font.Italic) styles.Add("斜體");
                if (run.Font.Underline != Underline.None) styles.Add($"底線({run.Font.Underline})");
                if (styles.Count > 0)
                    result.AppendLine($"    樣式: {string.Join(", ", styles)}");
            }
            
            if (para.Runs.Count > 10)
            {
                result.AppendLine($"\n  ... 還有 {para.Runs.Count - 10} 個 Run（已省略）");
            }
            result.AppendLine();
        }

        // JSON format for easy copying
        result.AppendLine("【JSON 格式（可用於 word_edit_paragraph）】");
        result.AppendLine("{");
        result.AppendLine($"  \"alignment\": \"{format.Alignment.ToString().ToLower()}\",");
        result.AppendLine($"  \"leftIndent\": {format.LeftIndent:F2},");
        result.AppendLine($"  \"rightIndent\": {format.RightIndent:F2},");
        result.AppendLine($"  \"firstLineIndent\": {format.FirstLineIndent:F2},");
        result.AppendLine($"  \"spaceBefore\": {format.SpaceBefore:F2},");
        result.AppendLine($"  \"spaceAfter\": {format.SpaceAfter:F2},");
        result.AppendLine($"  \"lineSpacing\": {format.LineSpacing:F2}");
        
        if (para.Runs.Count > 0)
        {
            var run = para.Runs[0];
            result.AppendLine($"  \"fontNameAscii\": \"{run.Font.NameAscii}\",");
            result.AppendLine($"  \"fontNameFarEast\": \"{run.Font.NameFarEast}\",");
            result.AppendLine($"  \"fontSize\": {run.Font.Size},");
            result.AppendLine($"  \"bold\": {run.Font.Bold.ToString().ToLower()},");
            result.AppendLine($"  \"italic\": {run.Font.Italic.ToString().ToLower()}");
        }
        
        result.AppendLine("}");

        return await Task.FromResult(result.ToString());
    }
}

