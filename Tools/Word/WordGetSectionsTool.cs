using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetSectionsTool : IAsposeTool
{
    public string Description => "Get all sections information from a Word document";

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
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided returns all sections)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var result = new StringBuilder();

        result.AppendLine("=== 文檔節資訊 ===\n");
        result.AppendLine($"總節數: {doc.Sections.Count}\n");

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"節索引 {sectionIndex.Value} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
            }
            
            var section = doc.Sections[sectionIndex.Value];
            AppendSectionInfo(result, section, sectionIndex.Value);
        }
        else
        {
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                var section = doc.Sections[i];
                AppendSectionInfo(result, section, i);
                if (i < doc.Sections.Count - 1)
                {
                    result.AppendLine();
                }
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private void AppendSectionInfo(StringBuilder result, Section section, int index)
    {
        result.AppendLine($"【節 {index}】");
        
        var pageSetup = section.PageSetup;
        result.AppendLine($"頁面設置:");
        result.AppendLine($"  紙張大小: {pageSetup.PaperSize}");
        result.AppendLine($"  方向: {pageSetup.Orientation}");
        result.AppendLine($"  上邊距: {pageSetup.TopMargin} 點");
        result.AppendLine($"  下邊距: {pageSetup.BottomMargin} 點");
        result.AppendLine($"  左邊距: {pageSetup.LeftMargin} 點");
        result.AppendLine($"  右邊距: {pageSetup.RightMargin} 點");
        result.AppendLine($"  頁眉距離: {pageSetup.HeaderDistance} 點");
        result.AppendLine($"  頁尾距離: {pageSetup.FooterDistance} 點");
        result.AppendLine($"  頁碼起始: {(pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber.ToString() : "繼承上一節")}");
        result.AppendLine($"  不同首頁: {pageSetup.DifferentFirstPageHeaderFooter}");
        result.AppendLine($"  不同奇偶頁: {pageSetup.OddAndEvenPagesHeaderFooter}");
        result.AppendLine($"  分欄數: {pageSetup.TextColumns.Count}");
        
        result.AppendLine();
        result.AppendLine($"內容統計:");
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);
        result.AppendLine($"  段落數: {paragraphs.Count}");
        result.AppendLine($"  表格數: {tables.Count}");
        result.AppendLine($"  形狀數: {shapes.Count}");
        
        result.AppendLine();
        result.AppendLine($"頁眉頁尾:");
        var headerCount = 0;
        var footerCount = 0;
        foreach (HeaderFooter hf in section.HeadersFooters)
        {
            if (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                hf.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                hf.HeaderFooterType == HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                     hf.HeaderFooterType == HeaderFooterType.FooterFirst ||
                     hf.HeaderFooterType == HeaderFooterType.FooterEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    footerCount++;
            }
        }
        result.AppendLine($"  頁眉數: {headerCount}");
        result.AppendLine($"  頁尾數: {footerCount}");
    }
}

