using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCopyParagraphFormatTool : IAsposeTool
{
    public string Description => "Copy paragraph format from one paragraph to another (format brush functionality)";

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
            sourceParagraphIndex = new
            {
                type = "number",
                description = "Source paragraph index (0-based) - format will be copied from this paragraph"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index (0-based) - format will be applied to this paragraph"
            }
        },
        required = new[] { "path", "sourceParagraphIndex", "targetParagraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sourceParagraphIndex = arguments?["sourceParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceParagraphIndex is required");
        var targetParagraphIndex = arguments?["targetParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetParagraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (sourceParagraphIndex < 0 || sourceParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"來源段落索引 {sourceParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"目標段落索引 {targetParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var sourcePara = paragraphs[sourceParagraphIndex] as Paragraph;
        var targetPara = paragraphs[targetParagraphIndex] as Paragraph;
        
        if (sourcePara == null || targetPara == null)
        {
            throw new InvalidOperationException("無法獲取段落");
        }
        
        // Copy paragraph format
        targetPara.ParagraphFormat.StyleName = sourcePara.ParagraphFormat.StyleName;
        targetPara.ParagraphFormat.Alignment = sourcePara.ParagraphFormat.Alignment;
        targetPara.ParagraphFormat.LeftIndent = sourcePara.ParagraphFormat.LeftIndent;
        targetPara.ParagraphFormat.RightIndent = sourcePara.ParagraphFormat.RightIndent;
        targetPara.ParagraphFormat.FirstLineIndent = sourcePara.ParagraphFormat.FirstLineIndent;
        targetPara.ParagraphFormat.SpaceBefore = sourcePara.ParagraphFormat.SpaceBefore;
        targetPara.ParagraphFormat.SpaceAfter = sourcePara.ParagraphFormat.SpaceAfter;
        targetPara.ParagraphFormat.LineSpacing = sourcePara.ParagraphFormat.LineSpacing;
        targetPara.ParagraphFormat.LineSpacingRule = sourcePara.ParagraphFormat.LineSpacingRule;
        
        // Copy tab stops
        targetPara.ParagraphFormat.TabStops.Clear();
        for (int i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功複製段落格式\n";
        result += $"來源段落: #{sourceParagraphIndex}\n";
        result += $"目標段落: #{targetParagraphIndex}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

