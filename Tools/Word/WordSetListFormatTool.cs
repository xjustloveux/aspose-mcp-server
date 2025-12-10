using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeMcpServer.Tools;

public class WordSetListFormatTool : IAsposeTool
{
    public string Description => "Set list format (number style, indentation) for existing list items in Word document";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index of the list item to format (0-based)"
            },
            numberStyle = new
            {
                type = "string",
                description = "Number style: arabic, roman, letter, bullet, none (optional)",
                @enum = new[] { "arabic", "roman", "letter", "bullet", "none" }
            },
            indentLevel = new
            {
                type = "number",
                description = "Indentation level (0-8, optional). Each level = 36 points (0.5 inch)"
            },
            leftIndent = new
            {
                type = "number",
                description = "Left indent in points (optional, overrides indentLevel if provided)"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indent in points (optional, negative for hanging indent)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var numberStyle = arguments?["numberStyle"]?.GetValue<string>();
        var indentLevel = arguments?["indentLevel"]?.GetValue<int?>();
        var leftIndent = arguments?["leftIndent"]?.GetValue<double?>();
        var firstLineIndent = arguments?["firstLineIndent"]?.GetValue<double?>();

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
        
        var changes = new List<string>();
        
        // Set number style if provided
        if (!string.IsNullOrEmpty(numberStyle) && para.ListFormat.IsListItem)
        {
            var list = para.ListFormat.List;
            if (list != null)
            {
                var level = para.ListFormat.ListLevelNumber;
                var listLevel = list.ListLevels[level];
                
                var style = numberStyle.ToLower() switch
                {
                    "arabic" => NumberStyle.Arabic,
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    "bullet" => NumberStyle.Bullet,
                    "none" => NumberStyle.None,
                    _ => NumberStyle.Arabic
                };
                
                listLevel.NumberStyle = style;
                changes.Add($"編號樣式: {numberStyle}");
            }
        }
        
        // Set indentation
        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = indentLevel.Value * 36; // Each level = 36 points
            changes.Add($"縮排層級: {indentLevel.Value}");
        }
        
        if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
            changes.Add($"左縮排: {leftIndent.Value} 點");
        }
        
        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
            changes.Add($"首行縮排: {firstLineIndent.Value} 點");
        }
        
        doc.Save(outputPath);
        
        var result = $"成功設定清單格式\n";
        result += $"段落索引: {paragraphIndex}\n";
        if (changes.Count > 0)
        {
            result += $"變更內容: {string.Join("、", changes)}\n";
        }
        else
        {
            result += "未提供變更參數\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

