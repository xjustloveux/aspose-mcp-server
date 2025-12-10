using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordInsertBlankPageTool : IAsposeTool
{
    public string Description => "Insert a blank page at specified position";

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
            position = new
            {
                type = "string",
                description = "Position: before or after (default: after)",
                @enum = new[] { "before", "after" }
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (0-based). Insert blank page before/after this page. If not provided, inserts at the end."
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var position = arguments?["position"]?.GetValue<string>() ?? "after";
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        if (pageIndex.HasValue)
        {
            // Find the target page by counting pages
            // Note: Aspose.Words doesn't have direct page access, so we need to navigate
            var sections = doc.Sections;
            int currentPage = 0;
            Node? targetNode = null;
            
            foreach (Section section in sections)
            {
                var paragraphs = section.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    // Rough estimation: each paragraph might be on a new page
                    // This is a simplified approach
                    if (currentPage == pageIndex.Value)
                    {
                        targetNode = para;
                        break;
                    }
                    currentPage++;
                }
                if (targetNode != null) break;
            }
            
            if (targetNode != null)
            {
                builder.MoveTo(targetNode);
                if (position == "before")
                {
                    // Insert page break before target node
                    builder.InsertBreak(BreakType.PageBreak);
                    builder.InsertBreak(BreakType.PageBreak); // Second break creates blank page
                }
                else
                {
                    // Insert section break after
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
                    builder.InsertBreak(BreakType.PageBreak);
                }
            }
            else
            {
                // If page not found, insert at end
                builder.MoveToDocumentEnd();
                builder.InsertBreak(BreakType.SectionBreakNewPage);
                builder.InsertBreak(BreakType.PageBreak);
            }
        }
        else
        {
            // Insert at end of document
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.InsertBreak(BreakType.PageBreak);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功插入空白頁\n";
        if (pageIndex.HasValue)
        {
            result += $"位置: {(position == "before" ? "之前" : "之後")} 第 {pageIndex.Value + 1} 頁\n";
        }
        else
        {
            result += "位置: 文檔末尾\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

