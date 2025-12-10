using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeMcpServer.Tools;

public class WordGetListFormatTool : IAsposeTool
{
    public string Description => "Get list format information for paragraphs in a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, optional, if not provided returns all list paragraphs)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        var result = new StringBuilder();

        result.AppendLine("=== 文檔列表格式資訊 ===\n");

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
            
            var para = paragraphs[paragraphIndex.Value];
            AppendListFormatInfo(result, para, paragraphIndex.Value);
        }
        else
        {
            var listParagraphs = paragraphs
                .Where(p => p.ListFormat != null && p.ListFormat.IsListItem)
                .ToList();
            
            result.AppendLine($"總列表段落數: {listParagraphs.Count}\n");
            
            if (listParagraphs.Count == 0)
            {
                result.AppendLine("未找到列表段落");
                return await Task.FromResult(result.ToString());
            }
            
            for (int i = 0; i < listParagraphs.Count; i++)
            {
                var para = listParagraphs[i];
                var paraIndex = paragraphs.IndexOf(para);
                AppendListFormatInfo(result, para, paraIndex);
                if (i < listParagraphs.Count - 1)
                {
                    result.AppendLine();
                }
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendListFormatInfo(StringBuilder result, Paragraph para, int paraIndex)
    {
        result.AppendLine($"【段落 {paraIndex}】");
        result.AppendLine($"內容預覽: {para.GetText().Trim().Substring(0, Math.Min(50, para.GetText().Trim().Length))}...");
        
        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine($"是否列表項: 是");
            result.AppendLine($"列表級別: {para.ListFormat.ListLevelNumber}");
            
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"列表ID: {para.ListFormat.List.ListId}");
            }
            
            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                result.AppendLine($"列表符號: {level.NumberFormat}");
                result.AppendLine($"對齊方式: {level.Alignment}");
                result.AppendLine($"文本位置: {level.TextPosition}");
                result.AppendLine($"編號樣式: {level.NumberStyle}");
            }
        }
        else
        {
            result.AppendLine($"是否列表項: 否");
        }
    }
}

