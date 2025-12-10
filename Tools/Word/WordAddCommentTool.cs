using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddCommentTool : IAsposeTool
{
    public string Description => "Add a comment to a Word document at specified paragraph or position";

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
            text = new
            {
                type = "string",
                description = "Comment text content"
            },
            author = new
            {
                type = "string",
                description = "Comment author name (optional, defaults to 'Comment Author')"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to attach comment to (0-based). If not provided, attaches to last paragraph."
            },
            startRunIndex = new
            {
                type = "number",
                description = "Start run index within the paragraph (0-based, optional)"
            },
            endRunIndex = new
            {
                type = "number",
                description = "End run index within the paragraph (0-based, optional). If not provided, comment covers the entire paragraph or specified run."
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var author = arguments?["author"]?.GetValue<string>() ?? "Comment Author";
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var startRunIndex = arguments?["startRunIndex"]?.GetValue<int?>();
        var endRunIndex = arguments?["endRunIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        Paragraph? targetPara = null;
        
        // Determine target paragraph
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
            targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
            if (targetPara == null)
            {
                throw new ArgumentException($"無法找到索引 {paragraphIndex.Value} 的段落");
            }
        }
        else
        {
            // Default: use last paragraph
            if (paragraphs.Count > 0)
            {
                targetPara = paragraphs[paragraphs.Count - 1] as Paragraph;
                if (targetPara == null)
                {
                    throw new InvalidOperationException("無法找到有效的段落");
                }
            }
            else
            {
                throw new InvalidOperationException("文檔中沒有段落可以添加註解");
            }
        }
        
        // Create comment
        var comment = new Comment(doc)
        {
            Author = author,
            Initial = author.Length >= 2 ? author.Substring(0, 2).ToUpper() : author.ToUpper(),
            DateTime = System.DateTime.Now
        };
        
        // Add comment text
        var commentPara = new Paragraph(doc);
        commentPara.Runs.Add(new Run(doc, text));
        comment.Paragraphs.Add(commentPara);
        
        // Determine comment range and insert comment
        if (startRunIndex.HasValue && endRunIndex.HasValue)
        {
            // Comment on specific runs
            var runs = targetPara.GetChildNodes(NodeType.Run, false)!;
            if (runs.Count == 0)
            {
                throw new InvalidOperationException("無法獲取 Run 節點");
            }
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count ||
                endRunIndex.Value < 0 || endRunIndex.Value >= runs.Count ||
                startRunIndex.Value > endRunIndex.Value)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            
            var startRun = runs[startRunIndex.Value] as Run;
            var endRun = runs[endRunIndex.Value] as Run;
            
            if (startRun != null && endRun != null && startRun.ParentNode != null && endRun.ParentNode != null)
            {
                var rangeStart = new CommentRangeStart(doc, comment.Id);
                var rangeEnd = new CommentRangeEnd(doc, comment.Id);
                
                startRun.ParentNode.InsertBefore(rangeStart, startRun);
                endRun.ParentNode.InsertAfter(rangeEnd, endRun);
            }
        }
        else if (startRunIndex.HasValue)
        {
            // Comment on single run
            var runs = targetPara.GetChildNodes(NodeType.Run, false)!;
            if (runs.Count == 0)
            {
                throw new InvalidOperationException("無法獲取 Run 節點");
            }
            if (startRunIndex.Value < 0 || startRunIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"Run 索引超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            
            var run = runs[startRunIndex.Value] as Run;
            if (run != null && run.ParentNode != null)
            {
                var rangeStart = new CommentRangeStart(doc, comment.Id);
                var rangeEnd = new CommentRangeEnd(doc, comment.Id);
                
                run.ParentNode.InsertBefore(rangeStart, run);
                run.ParentNode.InsertAfter(rangeEnd, run);
            }
        }
        else
        {
            // Comment on entire paragraph - insert at end of paragraph
            var rangeStart = new CommentRangeStart(doc, comment.Id);
            var rangeEnd = new CommentRangeEnd(doc, comment.Id);
            
            targetPara?.AppendChild(rangeStart);
            targetPara?.AppendChild(rangeEnd);
        }
        
        // Insert comment into document
        if (doc.FirstSection?.Body != null)
        {
            doc.FirstSection.Body.AppendChild(comment);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功添加註解\n";
        result += $"作者: {author}\n";
        result += $"內容: {text}\n";
        if (paragraphIndex.HasValue)
        {
            result += $"段落索引: {paragraphIndex.Value}\n";
        }
        if (startRunIndex.HasValue)
        {
            result += $"Run 範圍: {startRunIndex.Value}";
            if (endRunIndex.HasValue)
            {
                result += $" - {endRunIndex.Value}";
            }
            result += "\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

