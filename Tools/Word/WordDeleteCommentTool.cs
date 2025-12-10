using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteCommentTool : IAsposeTool
{
    public string Description => "Delete a specific comment from Word document";

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
            }
        },
        required = new[] { "path", "commentIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var commentIndex = arguments?["commentIndex"]?.GetValue<int>() ?? throw new ArgumentException("commentIndex is required");

        var doc = new Document(path);
        
        // Get all comments
        var comments = doc.GetChildNodes(NodeType.Comment, true);
        
        if (commentIndex < 0 || commentIndex >= comments.Count)
        {
            throw new ArgumentException($"註解索引 {commentIndex} 超出範圍 (文檔共有 {comments.Count} 個註解)");
        }
        
        var commentToDelete = comments[commentIndex] as Comment;
        if (commentToDelete == null)
        {
            throw new InvalidOperationException($"無法找到索引 {commentIndex} 的註解");
        }
        
        // Get comment info before deletion
        string author = commentToDelete.Author;
        string commentText = commentToDelete.GetText().Trim();
        string preview = commentText.Length > 50 ? commentText.Substring(0, 50) + "..." : commentText;
        
        // Remove comment range markers if they exist
        var rangeStarts = doc.GetChildNodes(NodeType.CommentRangeStart, true);
        var rangeEnds = doc.GetChildNodes(NodeType.CommentRangeEnd, true);
        
        foreach (CommentRangeStart rangeStart in rangeStarts)
        {
            if (rangeStart.Id == commentToDelete.Id)
            {
                rangeStart.Remove();
            }
        }
        
        foreach (CommentRangeEnd rangeEnd in rangeEnds)
        {
            if (rangeEnd.Id == commentToDelete.Id)
            {
                rangeEnd.Remove();
            }
        }
        
        // Delete the comment
        commentToDelete.Remove();
        
        doc.Save(outputPath);
        
        // Count remaining comments
        int remainingCount = doc.GetChildNodes(NodeType.Comment, true).Count;
        
        var result = $"成功刪除註解 #{commentIndex}\n";
        result += $"作者: {author}\n";
        result += $"內容預覽: {preview}\n";
        result += $"文檔剩餘註解數: {remainingCount}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

