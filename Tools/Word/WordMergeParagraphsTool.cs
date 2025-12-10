using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordMergeParagraphsTool : IAsposeTool
{
    public string Description => "Merge multiple paragraphs into one";

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
            startParagraphIndex = new
            {
                type = "number",
                description = "Start paragraph index (0-based, inclusive)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based, inclusive)"
            }
        },
        required = new[] { "path", "startParagraphIndex", "endParagraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"結束段落索引 {endParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (startParagraphIndex > endParagraphIndex)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 不能大於結束段落索引 {endParagraphIndex}");
        }
        
        if (startParagraphIndex == endParagraphIndex)
        {
            throw new ArgumentException("起始和結束段落索引相同，無需合併");
        }
        
        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        if (startPara == null)
        {
            throw new InvalidOperationException("無法獲取起始段落");
        }
        
        // Merge paragraphs: move all runs from subsequent paragraphs to the first paragraph
        for (int i = startParagraphIndex + 1; i <= endParagraphIndex; i++)
        {
            var para = paragraphs[i] as Paragraph;
            if (para != null)
            {
                // Add a space before merging if needed
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }
                
                // Move all runs from this paragraph to the start paragraph
                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove)
                {
                    startPara.AppendChild(run);
                }
                
                // Remove the merged paragraph
                para.Remove();
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功合併段落\n";
        result += $"合併範圍: 段落 #{startParagraphIndex} 到 #{endParagraphIndex}\n";
        result += $"合併段落數: {endParagraphIndex - startParagraphIndex + 1}\n";
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

