using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordReplyCommentTool : IAsposeTool
{
    public string Description => "Reply to an existing comment in Word document";

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
            commentIndex = new
            {
                type = "number",
                description = "Comment index (0-based, from word_get_comments)"
            },
            replyText = new
            {
                type = "string",
                description = "Reply text content"
            },
            author = new
            {
                type = "string",
                description = "Reply author name (optional, defaults to 'Reply Author')"
            }
        },
        required = new[] { "path", "commentIndex", "replyText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var commentIndex = arguments?["commentIndex"]?.GetValue<int>() ?? throw new ArgumentException("commentIndex is required");
        var replyText = arguments?["replyText"]?.GetValue<string>() ?? throw new ArgumentException("replyText is required");
        var author = arguments?["author"]?.GetValue<string>() ?? "Reply Author";

        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (commentIndex < 0 || commentIndex >= comments.Count)
        {
            throw new ArgumentException($"註解索引 {commentIndex} 超出範圍 (文檔共有 {comments.Count} 個註解)");
        }
        
        var parentComment = comments[commentIndex] as Comment;
        if (parentComment == null)
        {
            throw new InvalidOperationException($"無法找到索引 {commentIndex} 的註解");
        }
        
        // Create reply comment
        var replyComment = new Comment(doc)
        {
            Author = author,
            Initial = author.Substring(0, Math.Min(2, author.Length)).ToUpper()
        };
        
        replyComment.Paragraphs.Add(new Paragraph(doc));
        replyComment.FirstParagraph.Runs.Add(new Run(doc, replyText));
        
        // Add reply to parent comment
        // In Aspose.Words, replies are added as child comments
        parentComment.AppendChild(replyComment);
        
        doc.Save(outputPath);
        
        var result = $"成功回覆註解 #{commentIndex}\n";
        result += $"原註解作者: {parentComment.Author}\n";
        result += $"回覆作者: {author}\n";
        result += $"回覆內容: {replyText}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

