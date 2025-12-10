using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetDocumentInfoTool : IAsposeTool
{
    public string Description => "Get detailed document information including margins, compatibility mode, tab stops, and page setup";

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
            includeTabStops = new
            {
                type = "boolean",
                description = "Include tab stops information from header/footer (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeTabStops = arguments?["includeTabStops"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var result = new StringBuilder();
        
        result.AppendLine("=== Word 文件詳細資訊 ===\n");

        // Basic file info
        result.AppendLine("【檔案資訊】");
        result.AppendLine($"檔案格式: {doc.OriginalFileName}");
        result.AppendLine($"節數: {doc.Sections.Count}");
        result.AppendLine($"段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}");
        result.AppendLine($"表格數: {doc.GetChildNodes(NodeType.Table, true).Count}\n");

        // Compatibility mode
        result.AppendLine("【相容模式】");
        result.AppendLine("文件相容性設定:");
        result.AppendLine("  (使用 word_create 的 compatibilityMode 參數可設定為 Word2019)");
        result.AppendLine();

        // Page setup for first section
        var section = doc.FirstSection;
        if (section != null)
        {
            var pageSetup = section.PageSetup;
            
            result.AppendLine("【頁面設定】（第一節）");
            result.AppendLine($"頁面寬度: {pageSetup.PageWidth:F2} pt ({pageSetup.PageWidth / 28.35:F2} cm)");
            result.AppendLine($"頁面高度: {pageSetup.PageHeight:F2} pt ({pageSetup.PageHeight / 28.35:F2} cm)");
            result.AppendLine($"方向: {pageSetup.Orientation}");
            result.AppendLine();
            
            result.AppendLine("【邊界設定】");
            result.AppendLine($"上邊界: {pageSetup.TopMargin:F2} pt ({pageSetup.TopMargin / 28.35:F2} cm)");
            result.AppendLine($"下邊界: {pageSetup.BottomMargin:F2} pt ({pageSetup.BottomMargin / 28.35:F2} cm)");
            result.AppendLine($"左邊界: {pageSetup.LeftMargin:F2} pt ({pageSetup.LeftMargin / 28.35:F2} cm)");
            result.AppendLine($"右邊界: {pageSetup.RightMargin:F2} pt ({pageSetup.RightMargin / 28.35:F2} cm)");
            result.AppendLine();
            
            result.AppendLine("【頁首頁尾距離】");
            result.AppendLine($"頁首距離頁面頂端: {pageSetup.HeaderDistance:F2} pt ({pageSetup.HeaderDistance / 28.35:F2} cm)");
            result.AppendLine($"頁尾距離頁面底端: {pageSetup.FooterDistance:F2} pt ({pageSetup.FooterDistance / 28.35:F2} cm)");
            result.AppendLine();

            // Tab stops in header/footer
            if (includeTabStops)
            {
                result.AppendLine("【頁首 Tab 停駐點】");
                var headersFooters = section.HeadersFooters;
                var header = headersFooters[HeaderFooterType.HeaderPrimary];
                
                if (header != null && header.FirstParagraph != null)
                {
                    var tabStops = header.FirstParagraph.ParagraphFormat.TabStops;
                    if (tabStops.Count > 0)
                    {
                        for (int i = 0; i < tabStops.Count; i++)
                        {
                            var tab = tabStops[i];
                            result.AppendLine($"  Tab Stop {i + 1}:");
                            result.AppendLine($"    位置: {tab.Position:F2} pt ({tab.Position / 28.35:F2} cm)");
                            result.AppendLine($"    對齊: {tab.Alignment}");
                            result.AppendLine($"    前導字元: {tab.Leader}");
                        }
                    }
                    else
                    {
                        result.AppendLine("  無 Tab 停駐點");
                    }
                }
                else
                {
                    result.AppendLine("  無頁首");
                }
                result.AppendLine();

                result.AppendLine("【頁尾 Tab 停駐點】");
                var footer = headersFooters[HeaderFooterType.FooterPrimary];
                
                if (footer != null && footer.FirstParagraph != null)
                {
                    var tabStops = footer.FirstParagraph.ParagraphFormat.TabStops;
                    if (tabStops.Count > 0)
                    {
                        for (int i = 0; i < tabStops.Count; i++)
                        {
                            var tab = tabStops[i];
                            result.AppendLine($"  Tab Stop {i + 1}:");
                            result.AppendLine($"    位置: {tab.Position:F2} pt ({tab.Position / 28.35:F2} cm)");
                            result.AppendLine($"    對齊: {tab.Alignment}");
                            result.AppendLine($"    前導字元: {tab.Leader}");
                        }
                    }
                    else
                    {
                        result.AppendLine("  無 Tab 停駐點");
                    }
                }
                else
                {
                    result.AppendLine("  無頁尾");
                }
                result.AppendLine();
            }

            // Content width calculation
            var contentWidth = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
            result.AppendLine("【計算資訊】");
            result.AppendLine($"內容區域寬度: {contentWidth:F2} pt ({contentWidth / 28.35:F2} cm)");
            result.AppendLine($"中央位置（頁面中心）: {pageSetup.PageWidth / 2:F2} pt ({pageSetup.PageWidth / 2 / 28.35:F2} cm)");
            result.AppendLine($"右側位置（頁寬-右邊界）: {pageSetup.PageWidth - pageSetup.RightMargin:F2} pt ({(pageSetup.PageWidth - pageSetup.RightMargin) / 28.35:F2} cm)");
        }

        return await Task.FromResult(result.ToString());
    }
}

