using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteTextTool : IAsposeTool
{
    public string Description => "Delete text in specified range (by paragraph and run indices)";

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
                description = "Start paragraph index (0-based)"
            },
            startRunIndex = new
            {
                type = "number",
                description = "Start run index within start paragraph (0-based, optional, default: 0)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based, inclusive)"
            },
            endRunIndex = new
            {
                type = "number",
                description = "End run index within end paragraph (0-based, inclusive, optional, default: last run)"
            }
        },
        required = new[] { "path", "startParagraphIndex", "endParagraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var startRunIndex = arguments?["startRunIndex"]?.GetValue<int>() ?? 0;
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");
        var endRunIndex = arguments?["endRunIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count ||
            endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count ||
            startParagraphIndex > endParagraphIndex)
        {
            throw new ArgumentException($"段落索引超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        var endPara = paragraphs[endParagraphIndex] as Paragraph;
        
        if (startPara == null || endPara == null)
        {
            throw new InvalidOperationException("無法找到指定的段落");
        }
        
        // Get deleted text preview before deletion
        string deletedText = "";
        try
        {
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            var endRuns = endPara.GetChildNodes(NodeType.Run, false);
            
            if (startParagraphIndex == endParagraphIndex)
            {
                // Same paragraph
                if (startRuns != null && startRuns.Count > 0)
                {
                    var actualEndRunIndex = endRunIndex ?? (startRuns.Count - 1);
                    if (startRunIndex >= 0 && startRunIndex < startRuns.Count &&
                        actualEndRunIndex >= 0 && actualEndRunIndex < startRuns.Count &&
                        startRunIndex <= actualEndRunIndex)
                    {
                        for (int i = startRunIndex; i <= actualEndRunIndex; i++)
                        {
                            if (startRuns[i] is Run run)
                            {
                                deletedText += run.Text;
                            }
                        }
                    }
                }
            }
            else
            {
                // Multiple paragraphs
                if (startRuns != null && startRuns.Count > startRunIndex)
                {
                    for (int i = startRunIndex; i < startRuns.Count; i++)
                    {
                        if (startRuns[i] is Run run)
                        {
                            deletedText += run.Text;
                        }
                    }
                }
                
                // Middle paragraphs
                for (int p = startParagraphIndex + 1; p < endParagraphIndex; p++)
                {
                    var para = paragraphs[p] as Paragraph;
                    if (para != null)
                    {
                        deletedText += para.GetText();
                    }
                }
                
                // End paragraph
                if (endRuns != null && endRuns.Count > 0)
                {
                    var actualEndRunIndex = endRunIndex ?? (endRuns.Count - 1);
                    for (int i = 0; i <= actualEndRunIndex && i < endRuns.Count; i++)
                    {
                        if (endRuns[i] is Run run)
                        {
                            deletedText += run.Text;
                        }
                    }
                }
            }
        }
        catch
        {
            // Ignore preview errors
        }
        
        // Delete text
        if (startParagraphIndex == endParagraphIndex)
        {
            // Same paragraph - delete runs
            var runs = startPara.GetChildNodes(NodeType.Run, false);
            if (runs != null && runs.Count > 0)
            {
                var actualEndRunIndex = endRunIndex ?? (runs.Count - 1);
                if (startRunIndex >= 0 && startRunIndex < runs.Count &&
                    actualEndRunIndex >= 0 && actualEndRunIndex < runs.Count &&
                    startRunIndex <= actualEndRunIndex)
                {
                    // Delete from end to start to avoid index shifting
                    for (int i = actualEndRunIndex; i >= startRunIndex; i--)
                    {
                        runs[i]?.Remove();
                    }
                }
            }
        }
        else
        {
            // Multiple paragraphs - delete runs and paragraphs
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            if (startRuns != null && startRuns.Count > startRunIndex)
            {
                for (int i = startRuns.Count - 1; i >= startRunIndex; i--)
                {
                    startRuns[i]?.Remove();
                }
            }
            
            // Delete middle paragraphs
            for (int p = endParagraphIndex - 1; p > startParagraphIndex; p--)
            {
                paragraphs[p]?.Remove();
            }
            
            // Delete end paragraph runs
            var endRuns = endPara.GetChildNodes(NodeType.Run, false);
            if (endRuns != null && endRuns.Count > 0)
            {
                var actualEndRunIndex = endRunIndex ?? (endRuns.Count - 1);
                for (int i = actualEndRunIndex; i >= 0; i--)
                {
                    if (i < endRuns.Count)
                    {
                        endRuns[i]?.Remove();
                    }
                }
            }
        }
        
        doc.Save(outputPath);
        
        string preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;
        
        var result = $"成功刪除文字\n";
        result += $"範圍: 段落 {startParagraphIndex} Run {startRunIndex} 到 段落 {endParagraphIndex} Run {endRunIndex ?? -1}\n";
        if (!string.IsNullOrEmpty(preview))
        {
            result += $"刪除內容預覽: {preview}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

