using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetCommentsTool : IAsposeTool
{
    public string Description => "Get all comments from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (comments.Count == 0)
        {
            return await Task.FromResult("文檔中沒有找到註解");
        }
        
        var result = new System.Text.StringBuilder();
        result.AppendLine($"找到 {comments.Count} 個註解：\n");
        
        int index = 0;
        foreach (Comment comment in comments)
        {
            result.AppendLine($"註解 #{index}:");
            result.AppendLine($"  作者: {comment.Author}");
            result.AppendLine($"  初始: {comment.Initial}");
            result.AppendLine($"  日期: {comment.DateTime:yyyy-MM-dd HH:mm:ss}");
            
            // Get comment text
            string commentText = comment.GetText().Trim();
            if (commentText.Length > 100)
            {
                commentText = commentText.Substring(0, 100) + "...";
            }
            result.AppendLine($"  內容: {commentText}");
            
            // Get commented text range if available
            var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
            var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true);
            
            foreach (CommentRangeStart rangeStart in rangeStarts)
            {
                if (rangeStart.Id == comment.Id)
                {
                    result.AppendLine($"  範圍: 已標記文字");
                    break;
                }
            }
            
            result.AppendLine();
            index++;
        }
        
        return await Task.FromResult(result.ToString().TrimEnd());
    }
}

