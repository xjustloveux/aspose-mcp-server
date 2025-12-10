using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordEditListItemTool : IAsposeTool
{
    public string Description => "Edit a specific list item in Word document";

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
                description = "Paragraph index of the list item (0-based)"
            },
            text = new
            {
                type = "string",
                description = "New text content for the list item"
            },
            level = new
            {
                type = "number",
                description = "List level (0-8, optional). If provided, changes the indentation level."
            }
        },
        required = new[] { "path", "paragraphIndex", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var level = arguments?["level"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }
        
        // Replace text content
        para.Runs.Clear();
        var run = new Run(doc, text);
        para.AppendChild(run);
        
        // Update level if provided
        if (level.HasValue && level.Value >= 0 && level.Value <= 8)
        {
            para.ParagraphFormat.LeftIndent = level.Value * 36; // Each level = 36 points
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯清單項目\n";
        result += $"段落索引: {paragraphIndex}\n";
        result += $"新文字: {text}\n";
        if (level.HasValue)
        {
            result += $"級別: {level.Value}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

