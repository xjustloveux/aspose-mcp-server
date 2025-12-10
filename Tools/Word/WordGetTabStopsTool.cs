using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class WordGetTabStopsTool : IAsposeTool
{
    public string Description => "Get tab stops information from header, footer, or specific paragraph in a Word document";

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
            location = new
            {
                type = "string",
                description = "Where to get tab stops from: 'header', 'footer', or 'body' (default: body)",
                @enum = new[] { "header", "footer", "body" }
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based) if location='body'. Default: 0 (first paragraph)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0 (first section)"
            },
            allParagraphs = new
            {
                type = "boolean",
                description = "If true, read tab stops from all paragraphs (default: false, only first paragraph)"
            },
            includeStyle = new
            {
                type = "boolean",
                description = "If true, include tab stops from paragraph style (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var location = arguments?["location"]?.GetValue<string>() ?? "body";
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? 0;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var allParagraphs = arguments?["allParagraphs"]?.GetValue<bool>() ?? false;
        var includeStyle = arguments?["includeStyle"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var result = new StringBuilder();
        
        result.AppendLine($"=== Tab 停駐點資訊 ===");
        result.AppendLine($"位置: {location}");
        if (location == "body")
            result.AppendLine($"段落索引: {paragraphIndex}");
        result.AppendLine($"節索引: {sectionIndex}");
        result.AppendLine();

        // Get target paragraphs based on location
        List<Paragraph> targetParagraphs = new List<Paragraph>();
        string locationDesc = "";

        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header != null)
                {
                    var headerParas = header.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (allParagraphs)
                        targetParagraphs = headerParas;
                    else
                        targetParagraphs = headerParas.Count > 0 ? new List<Paragraph> { headerParas[0] } : new List<Paragraph>();
                    locationDesc = "頁首";
                }
                else
                {
                    throw new InvalidOperationException("找不到頁首");
                }
                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                {
                    var footerParas = footer.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                    if (allParagraphs)
                        targetParagraphs = footerParas;
                    else
                        targetParagraphs = footerParas.Count > 0 ? new List<Paragraph> { footerParas[0] } : new List<Paragraph>();
                    locationDesc = "頁尾";
                }
                else
                {
                    throw new InvalidOperationException("找不到頁尾");
                }
                break;

            case "body":
            default:
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                if (allParagraphs)
                    targetParagraphs = paragraphs;
                else
                {
                    if (paragraphIndex >= paragraphs.Count)
                        throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍（總段落數：{paragraphs.Count}）");
                    targetParagraphs = new List<Paragraph> { paragraphs[paragraphIndex] };
                }
                locationDesc = allParagraphs ? "內容" : $"內容段落 {paragraphIndex}";
                break;
        }

        if (targetParagraphs.Count == 0)
            throw new InvalidOperationException("找不到目標段落");

        // Collect all tab stops from all paragraphs and styles
        // Use position + alignment as key to handle cases where same position has different alignment types
        var allTabStops = new Dictionary<string, (double position, TabAlignment alignment, TabLeader leader, string source)>();
        
        for (int paraIdx = 0; paraIdx < targetParagraphs.Count; paraIdx++)
        {
            var para = targetParagraphs[paraIdx];
            var paraSource = allParagraphs ? $"段落 {paraIdx}" : "段落";
            
            // Get tab stops from paragraph itself
            var paraTabStops = para.ParagraphFormat.TabStops;
            for (int i = 0; i < paraTabStops.Count; i++)
            {
                var tab = paraTabStops[i];
                var position = Math.Round(tab.Position, 2); // Round to avoid floating point issues
                var key = $"{position}_{tab.Alignment}"; // Use position + alignment as unique key
                if (!allTabStops.ContainsKey(key))
                {
                    allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource}（自訂）");
                }
            }
            
            // Get tab stops from paragraph style if requested
            // Check style inheritance chain (style -> base style -> ...)
            if (includeStyle && para.ParagraphFormat.Style != null)
            {
                var paraDoc = para.Document;
                var currentStyle = para.ParagraphFormat.Style;
                var styleChain = new List<Style>();
                
                // Collect style inheritance chain
                while (currentStyle != null)
                {
                    styleChain.Add(currentStyle);
                    
                    // Get base style using BaseStyleName
                    if (!string.IsNullOrEmpty(currentStyle.BaseStyleName))
                    {
                        try
                        {
                            var baseStyle = paraDoc.Styles[currentStyle.BaseStyleName];
                            if (baseStyle != null && !styleChain.Contains(baseStyle))
                            {
                                currentStyle = baseStyle;
                            }
                            else
                            {
                                currentStyle = null; // Stop if base style not found or already in chain
                            }
                        }
                        catch
                        {
                            currentStyle = null; // Stop if error accessing base style
                        }
                    }
                    else
                    {
                        currentStyle = null; // No base style
                    }
                }
                
                // Read tab stops from all styles in the chain
                foreach (var chainStyle in styleChain)
                {
                    if (chainStyle.ParagraphFormat != null)
                    {
                        var styleTabStops = chainStyle.ParagraphFormat.TabStops;
                        for (int i = 0; i < styleTabStops.Count; i++)
                        {
                            var tab = styleTabStops[i];
                            var position = Math.Round(tab.Position, 2);
                            var key = $"{position}_{tab.Alignment}"; // Use position + alignment as unique key
                            
                            if (!allTabStops.ContainsKey(key))
                            {
                                var styleName = chainStyle == para.ParagraphFormat.Style 
                                    ? chainStyle.Name 
                                    : $"{para.ParagraphFormat.Style.Name}（基礎：{chainStyle.Name}）";
                                allTabStops[key] = (position, tab.Alignment, tab.Leader, $"{paraSource}（樣式：{styleName}）");
                            }
                        }
                    }
                }
            }
        }
        
        result.AppendLine($"【{locationDesc} 的 Tab 停駐點】");
        if (allParagraphs)
            result.AppendLine($"讀取範圍：所有段落（共 {targetParagraphs.Count} 個）");
        if (includeStyle)
            result.AppendLine($"包含樣式定位點：是");
        result.AppendLine();
        
        if (allTabStops.Count == 0)
        {
            result.AppendLine("  無 Tab 停駐點");
            result.AppendLine();
            result.AppendLine("說明：");
            result.AppendLine("  - 段落本身沒有自訂定位點");
            if (includeStyle)
                result.AppendLine("  - 樣式中也沒有定義定位點");
            result.AppendLine("  - 可能使用 Word 預設定位點（每 0.5 英吋）");
        }
        else
        {
            result.AppendLine($"  共 {allTabStops.Count} 個定位點：");
            result.AppendLine();
            
            int idx = 1;
            foreach (var kvp in allTabStops.OrderBy(x => x.Value.position))
            {
                var (position, alignment, leader, source) = kvp.Value;
                
                result.AppendLine($"  Tab Stop {idx}:");
                result.AppendLine($"    位置: {position:F2} pt ({position / 28.35:F2} cm) ({position / 12:F2} 字元)");
                result.AppendLine($"    對齊: {GetAlignmentName(alignment)}");
                result.AppendLine($"    前導字元: {GetLeaderName(leader)}");
                result.AppendLine($"    來源: {source}");
                
                // Add JSON format for easy copying
                var alignmentStr = alignment.ToString();
                result.AppendLine($"    JSON: {{\"position\": {position:F2}, \"alignment\": \"{alignmentStr}\"}}");
                result.AppendLine();
                idx++;
            }
        }
        
        // Add page setup info for reference
        result.AppendLine("【參考資訊】");
        result.AppendLine($"頁面寬度: {section.PageSetup.PageWidth:F2} pt ({section.PageSetup.PageWidth / 28.35:F2} cm)");
        result.AppendLine($"左邊界: {section.PageSetup.LeftMargin:F2} pt ({section.PageSetup.LeftMargin / 28.35:F2} cm)");
        result.AppendLine($"右邊界: {section.PageSetup.RightMargin:F2} pt ({section.PageSetup.RightMargin / 28.35:F2} cm)");
        result.AppendLine($"內容區域寬度: {section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin:F2} pt");
        result.AppendLine($"頁面中央位置: {section.PageSetup.PageWidth / 2:F2} pt");
        result.AppendLine($"右側位置（頁寬-右邊界）: {section.PageSetup.PageWidth - section.PageSetup.RightMargin:F2} pt");

        return await Task.FromResult(result.ToString());
    }

    private string GetAlignmentName(TabAlignment alignment)
    {
        return alignment switch
        {
            TabAlignment.Left => "Left（靠左）",
            TabAlignment.Center => "Center（置中）",
            TabAlignment.Right => "Right（靠右）",
            TabAlignment.Decimal => "Decimal（小數點）",
            TabAlignment.Bar => "Bar（列）",
            TabAlignment.Clear => "Clear（清除）",
            _ => alignment.ToString()
        };
    }

    private string GetLeaderName(TabLeader leader)
    {
        return leader switch
        {
            TabLeader.None => "None（無）",
            TabLeader.Dots => "Dots（點）",
            TabLeader.Dashes => "Dashes（虛線）",
            TabLeader.Line => "Line（底線）",
            TabLeader.Heavy => "Heavy（粗線）",
            TabLeader.MiddleDot => "MiddleDot（中點）",
            _ => leader.ToString()
        };
    }
}

