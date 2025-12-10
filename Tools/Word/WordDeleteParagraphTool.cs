using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteParagraphTool : IAsposeTool
{
    public string Description => "Delete a specific paragraph from Word document by index";

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
                description = "Index of the paragraph to delete (0-based). Use 0 to delete the first paragraph."
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");

        var doc = new Document(path);
        
        // Get all paragraphs in the document
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }

        var paragraphToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paragraphToDelete == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }

        // Get paragraph text preview before deletion (for confirmation message)
        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50)
        {
            textPreview = textPreview.Substring(0, 50) + "...";
        }
        
        // Delete the paragraph
        paragraphToDelete.Remove();

        doc.Save(outputPath);

        var result = $"成功刪除段落 #{paragraphIndex}\n";
        if (!string.IsNullOrEmpty(textPreview))
        {
            result += $"內容預覽: {textPreview}\n";
        }
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

