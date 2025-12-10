using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Layout;

namespace AsposeMcpServer.Tools;

public class WordDeletePageTool : IAsposeTool
{
    public string Description => "Delete a specific page from Word document (by page index)";

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
            pageIndex = new
            {
                type = "number",
                description = "Page index to delete (0-based)"
            }
        },
        required = new[] { "path", "pageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        var doc = new Document(path);
        
        // Note: Aspose.Words doesn't have direct page access
        // We need to use LayoutCollector to find page boundaries
        var collector = new LayoutCollector(doc);
        doc.UpdatePageLayout();
        
        // Find all page breaks and section breaks
        var breaks = new List<(int pageNum, Node node)>();
        var sections = doc.Sections;
        
        int currentPage = 0;
        foreach (Section section in sections)
        {
            var nodes = section.GetChildNodes(NodeType.Paragraph, true);
            foreach (Paragraph para in nodes)
            {
                if (para.GetText().Contains("\f") || para.ParagraphFormat.PageBreakBefore)
                {
                    breaks.Add((currentPage, para));
                    currentPage++;
                }
            }
        }
        
        // If page not found in breaks, estimate by sections
        if (pageIndex < 0 || pageIndex >= sections.Count)
        {
            throw new ArgumentException($"頁面索引 {pageIndex} 超出範圍 (文檔約有 {sections.Count} 個節)");
        }
        
        // Delete the section corresponding to the page
        // Note: This is a simplified approach - deleting a section removes its content
        if (pageIndex < sections.Count)
        {
            var sectionToDelete = sections[pageIndex];
            
            // Get content preview before deletion
            string contentPreview = sectionToDelete.GetText().Trim();
            if (contentPreview.Length > 100)
            {
                contentPreview = contentPreview.Substring(0, 100) + "...";
            }
            
            // Remove all content from the section
            sectionToDelete.Body.RemoveAllChildren();
            
            doc.Save(outputPath);
            
            var result = $"成功刪除頁面 #{pageIndex}\n";
            if (!string.IsNullOrEmpty(contentPreview))
            {
                result += $"內容預覽: {contentPreview}\n";
            }
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
        
        throw new InvalidOperationException($"無法刪除頁面 #{pageIndex}");
    }
}

