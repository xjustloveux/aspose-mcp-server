using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordSetPageNumberTool : IAsposeTool
{
    public string Description => "Set page number format and position in footer (fine-grained control)";

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
            format = new
            {
                type = "string",
                description = "Page number format: simple (Page X), total (Page X of Y), chinese (第X頁 共Y頁), custom (use template)",
                @enum = new[] { "simple", "total", "chinese", "custom" }
            },
            template = new
            {
                type = "string",
                description = "Custom template for page number (use {PAGE} for page number, {NUMPAGES} for total pages). Required if format='custom'"
            },
            position = new
            {
                type = "string",
                description = "Page number position: left, center, right (default: center)",
                @enum = new[] { "left", "center", "right" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to apply to all sections"
            },
            clearExisting = new
            {
                type = "boolean",
                description = "Clear existing page number before setting new one (default: false, append to existing content)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var format = arguments?["format"]?.GetValue<string>() ?? "simple";
        var template = arguments?["template"]?.GetValue<string>();
        var position = arguments?["position"]?.GetValue<string>() ?? "center";
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var clearExisting = arguments?["clearExisting"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : new[] { doc.Sections[sectionIndex] };

        foreach (Section section in sections)
        {
            var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                section.HeadersFooters.Add(footer);
            }

            // Clear existing page number fields if requested
            if (clearExisting)
            {
                var nodes = footer.GetChildNodes(NodeType.FieldStart, true).ToList();
                foreach (var node in nodes)
                {
                    if (node is FieldStart fieldStart)
                    {
                        if (fieldStart.FieldType == FieldType.FieldPage || fieldStart.FieldType == FieldType.FieldNumPages)
                        {
                            // Remove the entire field (FieldStart, FieldSeparator, FieldEnd, and content)
                            var currentNode = (Node)fieldStart;
                            while (currentNode != null)
                            {
                                var nextNode = currentNode.NextSibling;
                                if (currentNode.NodeType == NodeType.FieldEnd)
                                {
                                    currentNode.Remove();
                                    break;
                                }
                                currentNode.Remove();
                                currentNode = nextNode;
                            }
                        }
                    }
                }
            }

            // Create page number paragraph
            var pageNumPara = new Paragraph(doc);
            pageNumPara.ParagraphFormat.Alignment = GetAlignment(position);
            
            string pageNumText = "";
            if (format == "custom" && !string.IsNullOrEmpty(template))
            {
                pageNumText = template;
            }
            else
            {
                switch (format)
                {
                    case "simple":
                        pageNumText = "Page {PAGE}";
                        break;
                    case "total":
                        pageNumText = "Page {PAGE} of {NUMPAGES}";
                        break;
                    case "chinese":
                        pageNumText = "第{PAGE}頁 共{NUMPAGES}頁";
                        break;
                    default:
                        pageNumText = "Page {PAGE}";
                        break;
                }
            }

            // Insert field codes
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(pageNumPara);
            
            // Replace placeholders with actual field codes
            if (pageNumText.Contains("{PAGE}"))
            {
                var beforePage = pageNumText.Substring(0, pageNumText.IndexOf("{PAGE}"));
                var afterPage = pageNumText.Substring(pageNumText.IndexOf("{PAGE}") + 6);
                
                if (!string.IsNullOrEmpty(beforePage))
                    builder.Write(beforePage);
                
                builder.InsertField("PAGE", "");
                
                if (!string.IsNullOrEmpty(afterPage))
                {
                    if (afterPage.Contains("{NUMPAGES}"))
                    {
                        var beforeNumPages = afterPage.Substring(0, afterPage.IndexOf("{NUMPAGES}"));
                        var afterNumPages = afterPage.Substring(afterPage.IndexOf("{NUMPAGES}") + 10);
                        
                        if (!string.IsNullOrEmpty(beforeNumPages))
                            builder.Write(beforeNumPages);
                        
                        builder.InsertField("NUMPAGES", "");
                        
                        if (!string.IsNullOrEmpty(afterNumPages))
                            builder.Write(afterNumPages);
                    }
                    else
                    {
                        builder.Write(afterPage);
                    }
                }
            }
            else if (pageNumText.Contains("{NUMPAGES}"))
            {
                var beforeNumPages = pageNumText.Substring(0, pageNumText.IndexOf("{NUMPAGES}"));
                var afterNumPages = pageNumText.Substring(pageNumText.IndexOf("{NUMPAGES}") + 10);
                
                if (!string.IsNullOrEmpty(beforeNumPages))
                    builder.Write(beforeNumPages);
                
                builder.InsertField("NUMPAGES", "");
                
                if (!string.IsNullOrEmpty(afterNumPages))
                    builder.Write(afterNumPages);
            }
            else
            {
                builder.Write(pageNumText);
            }

            footer.AppendChild(pageNumPara);
        }

        doc.Save(outputPath);
        
        var sectionsDesc = sectionIndex == -1 ? "所有節" : $"第 {sectionIndex} 節";
        var formatDesc = format == "custom" ? $"自訂：{template}" : format;
        
        return await Task.FromResult($"成功設定頁碼（格式：{formatDesc}，位置：{position}）於 {sectionsDesc}");
    }
    
    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Center
        };
    }
}

